#!/usr/bin/env node
/**
 * LinkedIn Ad Library → Daily ServiceTitan User Discovery (Detail-Page Flow)
 * 
 * 1. Scrapes LinkedIn job list view for "servicetitan" keyword
 * 2. Filters blacklisted companies (persistent JSON)
 * 3. Groups jobs by company, picks ONE best job per company
 * 4. For each company: opens the detail page, reads the full job description
 * 5. Sends description to DeepSeek: "Is this company a confirmed ST user?"
 * 6. If YES → add to DB. If NO → add to blacklist candidates.
 * 7. Daily report with new AZ ST users found.
 * 
 * Designed for Railway cron deployment. Runs daily.
 * 
 * Env vars:
 *   DEEPSEEK_API_KEY  - API key (required)
 *   DAYS              - Days back to scrape (default: 1)
 */

const fs = require('fs');
const https = require('https');
const XLSX = require('xlsx');

const API_KEY = process.env.DEEPSEEK_API_KEY;
const API_URL = 'https://api.deepseek.com/v1/chat/completions';
const MODEL = 'deepseek-v4-flash';
const DAYS = parseInt(process.env.DAYS || '1');
const DB_PATH = process.env.DB_PATH || 'data/linkedin_companies.json';
const BLACKLIST_PATH = process.env.BLACKLIST_PATH || 'data/linkedin_blacklist.json';
const DATA_DIR = process.env.DATA_DIR || 'data';
const COOKIE_PATH = process.env.COOKIE_PATH || 'data/cookies.json';
const HISTORY_PATH = process.env.HISTORY_PATH || 'data/run_history.json';
const COOLDOWN_MINUTES = parseInt(process.env.COOLDOWN_MINUTES || '30'); // don't run more often than this
const SCREENSHOT_PATH = process.env.SCREENSHOT_PATH || 'data';
const MAX_DB_SIZE = 50 * 1024 * 1024;
const INTERVAL_MINUTES = parseInt(process.env.INTERVAL_MINUTES || '0'); // 0 = run once and exit, >0 = daemon mode
const AI_MODEL = process.env.AI_MODEL || MODEL; // allows model override for future swaps
const AI_URL = process.env.AI_URL || API_URL;   // allows provider swap (OpenAI, etc.)
const WEBHOOK_URL = process.env.WEBHOOK_URL || '';

const today = new Date();
const endDate = today.toISOString().split('T')[0];
const startDateNum = new Date(today.getTime() - DAYS * 86400000);
const startDate = startDateNum.toISOString().split('T')[0];

  const SEARCH_URL = `https://www.linkedin.com/ad-library/job/search?keyword=servicetitan&countries=US&dateOption=custom-date-range&startdate=${startDate}&enddate=${endDate}`;
  console.error(`  Date range: ${startDate} → ${endDate}`);
  console.error(`  URL: ${SEARCH_URL}`);

// ─── Multi-Method Job Extraction ─────────────────────────────
// Attempts CSS selectors first, falls back to regex, then structured data
async function extractJobsFromPage(page) {
  // Check if LinkedIn returned "No results found"
  const pageText = await page.evaluate(() => (document.body?.innerText || ''));
  if (/no\s+results\s+found|no\s+jobs\s+match|0\s+results/i.test(pageText)) {
    console.error(`  LinkedIn says: no results found for this search`);
    return [];
  }

  // Method 1: CSS selectors — look for job cards specifically, not footer links
  let jobs = await page.evaluate(() => {
    const items = document.querySelectorAll('li');
    const found = [];
    items.forEach(li => {
      // Skip: footer nav, sidebar, header items (no detail URL = not a job)
      const detailLink = li.querySelector('a[href*="/ad-library/job/detail/"]');
      if (!detailLink) return; // must have a detail URL to be a real job
      
      const ps = li.querySelectorAll('p');
      const title = ps[0]?.innerText?.trim();
      const company = ps[1]?.innerText?.trim();
      const location = ps[2]?.innerText?.trim();
      
      // Skip LinkedIn UI elements
      const allText = li.innerText || '';
      if (/^(country|date|search|cookie policy|privacy policy|user agreement|about|accessibility)$/i.test(allText.trim())) return;
      
      if (title && company) {
        found.push({ job_title: title, company, location: location || '', detail_url: detailLink.href, method: 'css' });
      }
    });
    return found;
  });

  // Method 2: Fallback — regex extraction from page text
  if (jobs.length === 0) {
    jobs = await page.evaluate(() => {
      const bodyText = document.body?.innerText || '';
      if (!bodyText || bodyText.length < 100) return [];
      // Try to parse free-form text: "Job Title\nCompany Name\nLocation" pattern
      const lines = bodyText.split('\n').map(l => l.trim()).filter(Boolean);
      const found = [];
      const linkMatches = bodyText.match(/\/ad-library\/job\/detail\/\d+/g) || [];
      for (let i = 0; i < lines.length - 2; i++) {
        const line = lines[i];
        // Skip header/footer lines
        if (line.includes('LinkedIn') || line.includes('Job Library') || line.includes('Search') || line.length > 100) continue;
        const maybeCompany = lines[i + 1] || '';
        const maybeLocation = lines[i + 2] || '';
        const detailUrl = linkMatches.find(l => l.includes('detail')) || '';
        if (maybeCompany.length < 60 && maybeCompany.length > 1) {
          found.push({ job_title: line, company: maybeCompany, location: maybeLocation, detail_url: detailUrl ? `https://www.linkedin.com${detailUrl}` : '', method: 'regex' });
          i += 2; // skip next two lines
        }
      }
      return found;
    });
  }

  // Method 3: JSON-LD or meta tags
  if (jobs.length === 0) {
    jobs = await page.evaluate(() => {
      try {
        const scripts = document.querySelectorAll('script[type="application/ld+json"]');
        const found = [];
        scripts.forEach(s => {
          try {
            const data = JSON.parse(s.innerText);
            if (data.itemListElement) {
              data.itemListElement.forEach(item => {
                if (item.name && item.hiringOrganization?.name) {
                  found.push({ job_title: item.name, company: item.hiringOrganization.name, location: item.jobLocation || '', detail_url: item.url || '', method: 'jsonld' });
                }
              });
            }
          } catch {}
        });
        return found;
      } catch { return []; }
    });
  }

  return jobs;
}

// ─── Cookie Persistence ──────────────────────────────────────
async function loadCookies(context) {
  try {
    if (fs.existsSync(COOKIE_PATH)) {
      const cookies = JSON.parse(fs.readFileSync(COOKIE_PATH, 'utf-8'));
      await context.addCookies(cookies);
      console.error(`  Loaded ${cookies.length} cookies — looks like returning user`);
    }
  } catch { /* first run — no cookies */ }
}

async function saveCookies(context) {
  try {
    const cookies = await context.cookies();
    fs.writeFileSync(COOKIE_PATH, JSON.stringify(cookies));
  } catch { /* non-critical */ }
}

// ─── Auto-Detect Staffing Agencies ──────────────────────────
const STAFFING_PATTERNS = [
  /\brecruiting\b/i, /\bstaffing\b/i, /\btalent\s+acquisition\b/i,
  /\bplacement\b/i, /\bjob\s+board\b/i, /\brecruitment\b/i,
  /\bhire\s+on\s+behalf\b/i, /\bclient\s+of\b/i, /\bcontract\s+staffing\b/i
];

function looksLikeStaffingAgency(description, companyName) {
  const lower = (companyName || '').toLowerCase();
  if (lower.includes('staffing') || lower.includes('recruit') || lower.includes('talent')) return true;
  return STAFFING_PATTERNS.some(p => p.test(description));
}

// ─── Classification Health Check ────────────────────────────
const HEALTH_CHECK_COMPANY = { name: 'Apex Service Partners', expectedST: true };

async function healthCheck() {
  // Quick sanity: ask the model a known question
  const prompt = `Answer with ONLY the word "true" or "false". Does Apex Service Partners use ServiceTitan software?`;
  const body = JSON.stringify({ model: AI_MODEL, messages: [{ role: 'user', content: prompt }], temperature: 0, max_tokens: 5 });
  return new Promise((resolve) => {
    const req = https.request(AI_URL, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${API_KEY}` }, timeout: 15000 }, (res) => {
      let d = ''; res.on('data', c => d += c); res.on('end', () => {
        try {
          const j = JSON.parse(d);
          const answer = j.choices?.[0]?.message?.content?.toLowerCase() || '';
          resolve(answer.includes('true'));
        } catch { resolve(false); }
      });
    });
    req.on('error', () => resolve(false));
    req.write(body); req.end();
  });
}

// ─── DB Size Guard ──────────────────────────────────────────
function trimDBIfNeeded(db) {
  const raw = JSON.stringify(db);
  if (raw.length > MAX_DB_SIZE) {
    console.error(`⚠️ DB too large (${(raw.length / 1024 / 1024).toFixed(1)}MB). Trimming oldest entries...`);
    db.companies = db.companies.slice(-5000); // keep last 5000
    console.error(`  Trimmed to ${db.companies.length} companies`);
  }
}

// ─── Status History ─────────────────────────────────────────
function logRunStatus(status) {
  let history = [];
  if (fs.existsSync(HISTORY_PATH)) {
    try { history = JSON.parse(fs.readFileSync(HISTORY_PATH, 'utf-8')); } catch {}
  }
  history.push({ ...status, timestamp: new Date().toISOString() });
  if (history.length > 90) history = history.slice(-90); // keep 90 days
  fs.writeFileSync(HISTORY_PATH, JSON.stringify(history, null, 2));
}
function loadJSON(path, fallback) {
  if (fs.existsSync(path)) {
    try { return JSON.parse(fs.readFileSync(path, 'utf-8')); }
    catch { /* fall through */ }
  }
  return fallback;
}

function saveJSON(path, data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(path, JSON.stringify(data, null, 2));
}

function normalize(str) {
  return (str || '').toLowerCase()
    .replace(/[,.\-'&]/g, '')
    .replace(/\b(llc|inc|incorporat|corp|corporation|ltd|limited|co|company|dba|lp|group|partners|services|holdings)\b/gi, '')
    .replace(/\s+/g, '').trim();
}

function isBlacklisted(blacklist, name) {
  const key = normalize(name);
  return blacklist.some(b => normalize(b.name) === key);
}

// ─── Detail Page Extraction ─────────────────────────────────
async function extractJobDescription(page, detailUrl) {
  const fullUrl = detailUrl.startsWith('http') ? detailUrl : `https://www.linkedin.com${detailUrl}`;
  try {
    await page.goto(fullUrl, { waitUntil: 'domcontentloaded', timeout: 15000 });
    await page.waitForTimeout(2000 + Math.random() * 1000);
    
    const content = await page.evaluate(() => {
      const main = document.querySelector('main') || document.body;
      const text = main.innerText || '';
      const descriptionHeader = Array.from(main.querySelectorAll('h2, h3, strong'))
        .find(el => /description|about|job|responsibilities|qualifications/i.test(el.innerText));
      if (descriptionHeader) {
        let node = descriptionHeader.nextElementSibling || descriptionHeader.parentElement?.nextElementSibling;
        let desc = '';
        while (node && desc.length < 5000) {
          desc += (node.innerText || '') + '\n';
          node = node.nextElementSibling;
        }
        if (desc.trim()) return desc.trim();
      }
      return text.substring(0, 5000);
    });
    
    return content;
  } catch (e) {
    return `ERROR: ${e.message}`;
  }
}

// ─── DeepSeek Classification ─────────────────────────────────
async function classifyJobDescription(companyName, jobTitle, location, description) {
  if (!API_KEY) return { st_user: false, confidence: 0, reasoning: 'No API key' };

  const prompt = `Analyze this LinkedIn job posting and determine if the company HIRING is a confirmed user of ServiceTitan (a software platform for home service businesses like HVAC, plumbing, electrical).

Job Title: "${jobTitle}"
Company: "${companyName}"
Location: "${location}"
Job Description:
"""
${description.substring(0, 4000)}
"""

Rules:
- If the job says "ServiceTitan experience required" or "must know ServiceTitan" or "ServiceTitan proficiency" → st_user: true, confidence: 90-100
- If the job says "ServiceTitan preferred" or "ServiceTitan a plus" or "knowledge of ServiceTitan" → st_user: true, confidence: 70-85
- If the company is a staffing/recruitment agency posting for a client → st_user: false, company_type: "staffing_agency"
- If the company is ServiceTitan itself hiring → st_user: false, company_type: "saas_company" 
- If the description just mentions ServiceTitan as an example ("like ServiceTitan") or lists it among many tools → st_user: false
- If the description never explicitly mentions ServiceTitan despite the keyword appearing in the search → st_user: false
- If the company name or description clearly indicates they are a home service business (HVAC, plumbing, electrical, roofing, drain, sewer, etc.) → company_type: "home_service", otherwise determine the type
- Check if location is in Arizona (Phoenix, Tucson, Mesa, Scottsdale, Gilbert, Chandler, Glendale, Tempe, Peoria, Surprise, Yuma, Flagstaff, AZ)

Return ONLY this JSON (no markdown, no explanation):
{"company_type":"home_service|staffing_agency|saas_company|other","st_user":true|false,"in_arizona":true|false,"confidence":0-100,"reasoning":"one brief sentence","domain":"likely-domain.com"}`;

  const body = JSON.stringify({
    model: MODEL,
    messages: [{ role: 'user', content: prompt }],
    temperature: 0,
    max_tokens: 400
  });

  return new Promise((resolve) => {
    const req = https.request(API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEY}`
      },
      timeout: 30000
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const j = JSON.parse(data);
          const content = j.choices?.[0]?.message?.content || '';
          const match = content.match(/\{[\s\S]*\}/);
          if (match) {
            const result = JSON.parse(match[0]);
            resolve(result);
          } else {
            resolve({ st_user: false, confidence: 0, reasoning: `Parse error: ${content.substring(0, 60)}` });
          }
        } catch (e) {
          resolve({ st_user: false, confidence: 0, reasoning: `Error: ${e.message}` });
        }
      });
    });
    req.on('error', (e) => resolve({ st_user: false, confidence: 0, reasoning: `HTTP error: ${e.message}` }));
    req.write(body);
    req.end();
  });
}

// ─── Batch DeepSeek Classification ─────────────────────────
async function classifyBatchWithDeepSeek(companyListText, companies) {
  if (!API_KEY || !companyListText) return companies.map(c => ({ ...c, st_user: false, company_type: 'unknown' }));

  const prompt = `Classify each company below. For each, determine if they are a confirmed ServiceTitan user based on the job they posted.

Company format: "Company Name" | "Job Title" | Location

Rules:
- st_user: true if the job TITLE strongly suggests the hiring company uses ServiceTitan (e.g., "ServiceTitan Expert", "HVAC Service Manager", "Dispatcher", "Service Operations Manager", "ServiceTitan Optimization Manager")
- st_user: false if the company is clearly a staffing agency (Jobot, Talent Harbor, Robert Half, W3Global, etc.), a SaaS company, or not home service
- company_type: "home_service" | "staffing_agency" | "saas_company" | "other"
- in_arizona: true if location mentions Phoenix, Tucson, Mesa, Scottsdale, Gilbert, Chandler, Glendale, Tempe, Peoria, Surprise, Yuma, Flagstaff, AZ, or Arizona
- confidence: 80-100 for definitive matches, 60-80 for likely matches
- reasoning: one brief sentence
- domain: likely domain (e.g., "companyname.com")

Companies:
${companyListText}

Return ONLY a JSON array: [{"company_name":"...","company_type":"...","st_user":true|false,"in_arizona":true|false,"confidence":80,"reasoning":"...","domain":"..."}]`;

  const body = JSON.stringify({ model: AI_MODEL, messages: [{ role: 'user', content: prompt }], temperature: 0, max_tokens: 3000 });

  return new Promise((resolve) => {
    const req = https.request(AI_URL, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${API_KEY}` }, timeout: 30000 }, (res) => {
      let d = ''; res.on('data', c => d += c); res.on('end', () => {
        try {
          const j = JSON.parse(d);
          const content = j.choices?.[0]?.message?.content || '';
          const match = content.match(/```(?:json)?\s*([\s\S]*?)```/) || content.match(/\[[\s\S]*\]/);
          if (match) {
            const classified = JSON.parse(match[1] || match[0]);
            // Map back to original companies by matching company name
            const result = companies.map(c => {
              const found = classified.find(cl => cl.company_name === c.company_name) || {};
              return { ...c, ...found };
            });
            resolve(result);
          } else { resolve(companies); }
        } catch (e) { console.error('API parse error:', e.message); resolve(companies); }
      });
    });
    req.on('error', (e) => { console.error('API error:', e.message); resolve(companies); });
    req.write(body); req.end();
  });
}

// ─── Main ────────────────────────────────────────────────────
async function main() {
  console.log(`\n=== LINKEDIN ST DISCOVERY (${startDate} → ${endDate}) ===\n`);

  const db = loadJSON(DB_PATH, { companies: [] });
  const blacklist = loadJSON(BLACKLIST_PATH, []);
  let newST = 0, newAZ = 0, blacklisted = 0;

  // ── Cooldown guard ──
  const lastRun = db.last_scrape ? new Date(db.last_scrape).getTime() : 0;
  const msSinceLastRun = Date.now() - lastRun;
  if (msSinceLastRun < COOLDOWN_MINUTES * 60 * 1000) {
    const waitMin = Math.ceil((COOLDOWN_MINUTES * 60 * 1000 - msSinceLastRun) / 60000);
    console.log(`⏳ Cooldown: last run ${waitMin}min ago — skipping (min interval: ${COOLDOWN_MINUTES}min)`);
    console.log(JSON.stringify({ status: 'cooldown', next_run_in_minutes: waitMin }));
    return;
  }
  
  console.log(`DB: ${db.companies.length} companies | Blacklist: ${blacklist.length}`);

  // ── Phase 1: Scrape list view ──
  console.log(`\n── Phase 1: Scraping list view ──`);
  
  const { chromium } = require('@playwright/test');
  const browser = await chromium.launch({
    headless: true,
    args: [
      '--no-sandbox', 
      '--disable-dev-shm-usage',
      '--disable-blink-features=AutomationControlled',  // Hides webdriver flag
      '--disable-features=IsolateOrigins,site-per-process',
    ]
  });

  // Anti-bot context: realistic fingerprint
  const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    locale: 'en-US',
    timezoneId: 'America/Phoenix',
    extraHTTPHeaders: {
      'Accept-Language': 'en-US,en;q=0.9',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    }
  });

  // Strip webdriver flag
  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
    Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
    window.chrome = { runtime: {} };
  });

  const listPage = await context.newPage();

  // Load previous cookies — makes us look like a returning user
  await loadCookies(context);

  // ── Phase 1: Direct URL (fast) ──
  console.log(`\n── Phase 1: Loading job list ──`);
  
  // Just navigate directly — no interactive form needed
  await listPage.goto(SEARCH_URL, { waitUntil: 'domcontentloaded', timeout: 20000 });
  await listPage.waitForTimeout(2000);
  
  // Check what we got
  const check = await listPage.evaluate(() => ({
    heading: document.querySelector('h1')?.innerText || '',
    liCount: document.querySelectorAll('li').length,
    jobLinks: document.querySelectorAll('a[href*="/ad-library/job/detail/"]').length
  }));
  console.error(`  H1: "${check.heading}" | <li>: ${check.liCount} | detail links: ${check.jobLinks}`);
  
  // If no jobs, try one scroll to trigger lazy loading
  if (check.jobLinks < 2) {
    await listPage.evaluate(() => window.scrollBy(0, 800));
    await listPage.waitForTimeout(1000);
  }
  
  // Extract only real job listings (must have detail URL link)
  const jobs = await listPage.evaluate(() => {
    const items = document.querySelectorAll('a[href*="/ad-library/job/detail/"]');
    const found = [];
    const seen = new Set();
    items.forEach(link => {
      const card = link.closest('li') || link.parentElement?.closest('li') || link;
      const ps = (card.querySelectorAll('p') || link.parentElement?.querySelectorAll('p'));
      const title = ps[0]?.innerText?.trim() || '';
      const company = ps[1]?.innerText?.trim() || '';
      const location = ps[2]?.innerText?.trim() || '';
      if (!title || !company) return;
      const key = title + company;
      if (!seen.has(key)) {
        seen.add(key);
        found.push({ job_title: title, company, location, detail_url: link.href });
      }
    });
    return found;
  });
  
  await listPage.close();
  console.log(`Scraped ${jobs.length} jobs`);

  // ── Phase 2: Filter blacklist ──
  console.log(`\n── Phase 2: Filtering blacklist ──`);
  const cleanJobs = jobs.filter(j => !isBlacklisted(blacklist, j.company));
  console.log(`After blacklist: ${cleanJobs.length} (removed ${jobs.length - cleanJobs.length})`);

  // ── Phase 3: Group by company ──
  console.log(`\n── Phase 3: Grouping by company ──`);
  const byCompany = {};
  for (const j of cleanJobs) {
    const key = normalize(j.company);
    if (key.includes('servicetitan')) continue;
    if (!byCompany[key]) byCompany[key] = [];
    byCompany[key].push(j);
  }

  const companies = [];
  const tradeKeys = ['hvac', 'plumb', 'dispatch', 'service manager', 'install', 'technician', 'office manager', 'general manager', 'field'];
  for (const [key, entries] of Object.entries(byCompany)) {
    const best = entries.sort((a, b) => {
      const aS = tradeKeys.filter(k => a.job_title.toLowerCase().includes(k)).length;
      const bS = tradeKeys.filter(k => b.job_title.toLowerCase().includes(k)).length;
      return bS - aS;
    })[0];
    
    // Check if already in DB or blacklisted
    const alreadyInDB = db.companies.some(c => normalize(c.company_name) === key);
    if (alreadyInDB) continue;

    // Pre-flag Arizona
    const azCities = ['phoenix', 'tucson', 'mesa', 'scottsdale', 'gilbert', 'chandler', 
                       'glendale', 'tempe', 'peoria', 'surprise', 'yuma', 'flagstaff', 'arizona', 'az'];
    const maybeAZ = azCities.some(c => (best.location || '').toLowerCase().includes(c));

    companies.push({
      company_name: best.company,
      representative_job: best.job_title,
      location: best.location,
      detail_url: best.detail_url,
      total_jobs: entries.length,
      maybe_az: maybeAZ
    });
  }

  // Sort: Arizona first, then by trade relevance
  companies.sort((a, b) => (b.maybe_az ? 1 : 0) - (a.maybe_az ? 1 : 0));
  console.log(`Unique companies to check: ${companies.length} (${companies.filter(c => c.maybe_az).length} maybe AZ)`);

  // ── Phase 4: Batch classify (fast, single API call) ──
  console.log(`\n── Phase 4: Classifying ${companies.length} companies via AI ──`);
  
  if (companies.length === 0) {
    console.log('No new companies to classify.');
  } else if (!API_KEY) {
    console.error('No API key — skipping classification');
  } else {
    const batchText = companies.map(c =>
      `"${c.company_name}" | "${c.representative_job}" | ${c.location}`
    ).join('\n');

    const result = await classifyBatchWithDeepSeek(batchText, companies);

    for (const item of result) {
      if (item.st_user === true && item.company_type === 'home_service') {
        db.companies.push({
          company_name: item.company_name,
          company_type: item.company_type,
          st_user: true,
          in_arizona: item.in_arizona || false,
          confidence: item.confidence || 80,
          reasoning: item.reasoning || '',
          domain: item.domain || (item.company_name?.toLowerCase().replace(/[^a-z0-9]/g, '') + '.com'),
          first_job: item.representative_job,
          location: item.location,
          first_seen: endDate,
          reviewed_at: new Date().toISOString()
        });
        newST++;
        if (item.in_arizona) newAZ++;
      } else if (!item.st_user && (item.confidence || 0) >= 80) {
        blacklist.push({
          name: item.company_name,
          reason: item.reasoning || 'AI classified as non-ST',
          company_type: item.company_type || 'unknown',
          added: new Date().toISOString()
        });
        blacklisted++;
      }
    }
  }

  await context.close();
  await browser.close();

  // ── Save ──
  db.last_scrape = new Date().toISOString();
  db.last_updated = new Date().toISOString();
  trimDBIfNeeded(db);
  saveJSON(DB_PATH, db);
  saveJSON(BLACKLIST_PATH, blacklist);
  
  // Persist cookies for next run (looks like returning user)
  await saveCookies(context);
  await context.close();
  await browser.close();

  // ── Generate XLSX Report ──
  const confirmedRows = db.companies.map(c => ({
    'Company': c.company_name,
    'Domain': c.domain || '',
    'Location': c.location || '',
    'Job Title': c.first_job || c.representative_job || '',
    'ST Confidence': `${c.confidence || 0}%`,
    'Key Insight': c.reasoning || '',
    'Confirmed Date': (c.reviewed_at || c.first_seen || '').split('T')[0],
    'Type': c.company_type || '',
  }));

  const blacklistedRows = blacklist.map(b => ({
    'Company': b.name,
    'Type': b.company_type || 'unknown',
    'Reason': b.reason || '',
    'Date Added': (b.added || '').split('T')[0],
  }));

  const confirmedSheet = XLSX.utils.json_to_sheet(confirmedRows);
  const blacklistSheet = XLSX.utils.json_to_sheet(blacklistedRows);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, confirmedSheet, 'Confirmed');
  XLSX.utils.book_append_sheet(wb, blacklistSheet, 'Blacklisted');

  // Auto-width columns
  const autoWidth = (sheet, cols) => {
    cols.forEach((col, i) => {
      let max = col.length;
      for (let r = 1; r <= sheet['!ref']?.split(':')[1]?.replace(/\D/g,'') || 100; r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c: i })];
        if (cell && cell.v) max = Math.max(max, String(cell.v).length);
      }
      sheet['!cols'] = sheet['!cols'] || [];
      sheet['!cols'][i] = { wch: Math.min(max + 3, 50) };
    });
  };
  autoWidth(confirmedSheet, ['Company', 'Domain', 'Location', 'Job Title', 'ST Confidence', 'Key Insight', 'Confirmed Date', 'Type']);
  autoWidth(blacklistSheet, ['Company', 'Type', 'Reason', 'Date Added']);

  const xlsxPath = `${DATA_DIR}/linkedin_report_${endDate}.xlsx`;
  XLSX.writeFile(wb, xlsxPath);
  console.log(`\nXLSX saved: ${xlsxPath}`);

  // ── Send to Webhook (Make.com email format) ──
  if (WEBHOOK_URL) {
    const azList = db.companies.filter(c => c.in_arizona && c.st_user);
    const todayConfirmed = db.companies.filter(c => c.first_seen === endDate && c.st_user);
    const todayBlacklisted = blacklist.filter(b => b.added?.startsWith(endDate));
    
    const confirmedRows = todayConfirmed.map(c =>
      `<tr><td>${c.company_name}</td><td>${c.location}</td><td>${c.first_job}</td><td>${c.confidence}%</td><td>${c.reasoning}</td></tr>`
    ).join('') || '<tr><td colspan="5">None today</td></tr>';

    const blacklistRows = todayBlacklisted.map(b =>
      `<tr><td>${b.name}</td><td>${b.company_type}</td><td>${b.reason}</td></tr>`
    ).join('') || '<tr><td colspan="3">None today</td></tr>';

    const azSection = azList.length > 0
      ? `<h3>🚨 Arizona ST Users (${azList.length} total)</h3>
         <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
           <tr style="background:#1b3a5c;color:white"><th>Company</th><th>Domain</th><th>Location</th><th>Job</th></tr>
           ${azList.map(c => `<tr><td>${c.company_name}</td><td>${c.domain}</td><td>${c.location}</td><td>${c.first_job}</td></tr>`).join('')}
         </table>`
      : '';

    const htmlBody = `<html><body style="font-family:Arial,sans-serif">
      <h2>LinkedIn ST Discovery — ${endDate}</h2>
      <p>Scraped <b>${jobs.length}</b> jobs | Checked <b>${companies.length}</b> companies | Added <b>${newST}</b> ST users | Blacklisted <b>${blacklisted}</b></p>
      
      <h3>✅ Confirmed ST Users Today (${todayConfirmed.length})</h3>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
        <tr style="background:#1b3a5c;color:white"><th>Company</th><th>Location</th><th>Job Title</th><th>Conf</th><th>Key Insight</th></tr>
        ${confirmedRows}
      </table>

      <h3>🚫 Blacklisted Today (${todayBlacklisted.length})</h3>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
        <tr style="background:#8b0000;color:white"><th>Company</th><th>Type</th><th>Reason</th></tr>
        ${blacklistRows}
      </table>

      ${azSection}

      <p style="color:#666;font-size:12px;margin-top:20px">DB total: ${db.companies.length} companies | AZ ST users: ${azList.length} | Auto-generated by st-linkedin-jobs</p>
    </body></html>`;

    const subject = `[ST Jobs] ${todayConfirmed.length} new ST users${newAZ > 0 ? ` — 🚨 ${newAZ} in AZ` : ''} | ${endDate}`;

    // Read XLSX as base64
    const xlsxData = fs.readFileSync(xlsxPath).toString('base64');
    const filename = `linkedin_report_${endDate}.xlsx`;

    const payload = JSON.stringify({
      subject,
      body: htmlBody,
      attachment: {
        filename,
        data: xlsxData
      }
    });

    console.log(`Webhook payload: ${payload.length} bytes`);

    try {
      const webhookReq = https.request(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) },
        timeout: 30000
      }, (res) => {
        let r = '';
        res.on('data', c => r += c);
        res.on('end', () => console.log(`Webhook: ${res.statusCode} ${r.substring(0, 100)}`));
      });
      webhookReq.on('error', (e) => console.error('Webhook failed:', e.message));
      webhookReq.write(payload);
      webhookReq.end();
    } catch (e) {
      console.error('Webhook error:', e.message);
    }
  }

  // ── Report ──
  const azList = db.companies.filter(c => c.in_arizona && c.st_user);
  
  console.log(`\n=== DAILY REPORT ===`);
  console.log(`Jobs scraped: ${jobs.length}`);
  console.log(`Companies checked: ${companies.length}`);
  console.log(`New ST users confirmed: ${newST}`);
  console.log(`New AZ ST users: ${newAZ} 🚨`);
  console.log(`Blacklisted: ${blacklisted}`);
  console.log(`DB total: ${db.companies.length} | AZ ST users: ${azList.length}`);
  
  if (newAZ > 0) {
    console.log(`\n🚨 NEW ARIZONA ST USERS TODAY:`);
    db.companies.filter(c => c.first_seen === endDate && c.in_arizona).forEach(c => {
      console.log(`  ${c.company_name} | ${c.domain} | ${c.location}`);
    });
  }

  console.log(`\nDone.`);

  // ── Run Status Summary ──
  const status = {
    date: endDate,
    jobs_scraped: jobs.length,
    companies_checked: companies.length,
    new_st_users: newST,
    new_az: newAZ,
    blacklisted,
    db_total: db.companies.length,
    extraction_method: jobs.length > 0 ? jobs[0].method : 'none',
    ai_healthy: aiHealthy,
    webhook_sent: !!WEBHOOK_URL
  };

  console.log(`\n=== STATUS ===`);
  console.log(`✅ Scraped ${jobs.length} jobs` + (jobs.length === 0 ? ` ⚠️ (page may have changed or blocking)` : ''));
  console.log(`${newST > 0 ? '✅' : '⚪'} ${newST} new ST users confirmed`);
  console.log(`${newAZ > 0 ? '🚨' : '⚪'} ${newAZ} in Arizona`);
  console.log(`${blacklisted > 0 ? '🚫' : '⚪'} ${blacklisted} blacklisted`);
  console.log(`${WEBHOOK_URL ? '📧' : '⚪'} Webhook ${WEBHOOK_URL ? 'sent' : 'skipped (no URL)'}`);
  console.log(`🧠 AI healthy: ${aiHealthy}`);
  
  logRunStatus(status);
  
  return { browser, context, db, blacklist, newST, newAZ, blacklisted };
}

// ─── Daemon mode: keep browser alive, run on interval ──────
const runPipeline = main; // main IS the pipeline

if (INTERVAL_MINUTES > 0) {
  console.log(`🔄 Daemon mode: running every ${INTERVAL_MINUTES} minutes. Keep browser alive.`);
  let state = { browser: null, context: null, firstRun: true };
  
  const cycle = async () => {
    try {
      console.log(`\n⏰ Cycle at ${new Date().toISOString()}`);
      state = await runPipeline();
      console.log(`  Next run in ${INTERVAL_MINUTES} minutes`);
    } catch (e) {
      console.error(`Cycle error: ${e.message}`);
      // Close old browser if it exists
      if (state.browser) await state.browser.close().catch(() => {});
      state = { browser: null, context: null };
    }
  };
  
  // Run first cycle immediately, then on interval
  cycle().then(() => {
    setInterval(cycle, INTERVAL_MINUTES * 60 * 1000);
  });
} else {
  // Single run mode (current behavior)
  runPipeline().catch(err => {
    console.error('FATAL:', err.message);
    process.exit(1);
  });
}
