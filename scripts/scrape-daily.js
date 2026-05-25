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
const WEBHOOK_URL = process.env.WEBHOOK_URL || '';

const today = new Date();
const endDate = today.toISOString().split('T')[0];
const startDateNum = new Date(today.getTime() - DAYS * 86400000);
const startDate = startDateNum.toISOString().split('T')[0];

const SEARCH_URL = `https://www.linkedin.com/ad-library/job/search?keyword=servicetitan&countries=US&dateOption=custom-date-range&startdate=${startDate}&enddate=${endDate}`;

// ─── DB & Blacklist ──────────────────────────────────────────
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

// ─── Main ────────────────────────────────────────────────────
async function main() {
  console.log(`\n=== LINKEDIN ST DISCOVERY (${startDate} → ${endDate}) ===\n`);

  const db = loadJSON(DB_PATH, { companies: [] });
  const blacklist = loadJSON(BLACKLIST_PATH, []);
  
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

  // Navigate with domcontentloaded (faster, less likely to timeout than networkidle)
  await listPage.goto('https://www.linkedin.com', { waitUntil: 'domcontentloaded', timeout: 15000 });
  await listPage.waitForTimeout(2000 + Math.random() * 2000);

  // Navigate to actual search — looks like a human clicking through
  await listPage.goto(SEARCH_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await listPage.waitForTimeout(3000 + Math.random() * 2000);

  const jobs = [];
  const seen = new Set();
  let prevCount = 0, noNew = 0;

  for (let i = 0; i < 80; i++) {
    const newJobs = await listPage.evaluate(() => {
      const items = document.querySelectorAll('li');
      const found = [];
      items.forEach(li => {
        const ps = li.querySelectorAll('p');
        const title = ps[0]?.innerText?.trim();
        const company = ps[1]?.innerText?.trim();
        const location = ps[2]?.innerText?.trim();
        const link = li.querySelector('a[href*="/ad-library/job/detail/"]')?.href;
        if (title && company) found.push({ job_title: title, company, location: location || '', detail_url: link || '' });
      });
      return found;
    });

    for (const j of newJobs) {
      const key = j.job_title + j.company;
      if (!seen.has(key)) { seen.add(key); jobs.push(j); }
    }

    // Human-like scroll: random distance, random delay
    const scrollAmount = 300 + Math.random() * 700;
    await listPage.evaluate((s) => window.scrollBy(0, s), scrollAmount);
    await listPage.waitForTimeout(1200 + Math.random() * 1500);

    if (jobs.length === prevCount) { noNew++; if (noNew >= 5) break; }
    else { noNew = 0; }
    prevCount = jobs.length;
  }

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

  // ── Phase 4: Open detail pages & classify ──
  console.log(`\n── Phase 4: Opening detail pages & classifying ──`);
  
  let newST = 0, newAZ = 0, blacklisted = 0;
  const classified = [];

  for (let i = 0; i < companies.length; i++) {
    const c = companies[i];
    if (!c.detail_url) {
      console.log(`[${i + 1}/${companies.length}] ${c.company_name} — NO DETAIL URL, skipping`);
      continue;
    }

    console.log(`[${i + 1}/${companies.length}] ${c.company_name} | "${c.representative_job}" | ${c.location}`);
    
    // Open detail page using same context (maintains cookies, looks natural)
    const detailPage = await context.newPage();
    const description = await extractJobDescription(detailPage, c.detail_url);
    await detailPage.close();

    if (!description || description.startsWith('ERROR')) {
      console.log(`  → SKIP: ${description || 'no description'}`);
      classified.push({ ...c, st_user: false, company_type: 'unknown', reason: 'no_description' });
      continue;
    }

    console.log(`  → Description: ${description.substring(0, 100)}...`);

    // Classify via DeepSeek
    const result = await classifyJobDescription(c.company_name, c.representative_job, c.location, description);
    
    const stUser = result.st_user === true;
    const inAZ = result.in_arizona === true;
    const companyType = result.company_type || 'unknown';
    const confidence = result.confidence || 0;
    const domain = result.domain || (c.company_name?.toLowerCase().replace(/[^a-z0-9]/g, '') + '.com');

    console.log(`  → ST: ${stUser} | Type: ${companyType} | AZ: ${inAZ} | ${confidence}% | ${result.reasoning}`);

    // If confirmed ST user → add to DB
    if (stUser && companyType === 'home_service') {
      db.companies.push({
        company_name: c.company_name,
        company_type: companyType,
        st_user: true,
        in_arizona: inAZ,
        confidence,
        reasoning: result.reasoning,
        domain,
        first_job: c.representative_job,
        location: c.location,
        first_seen: endDate,
        reviewed_at: new Date().toISOString()
      });
      newST++;
      if (inAZ) newAZ++;
    } 
    // If definitely NOT a ST user (staffing agency, other, false with high confidence)
    else if (!stUser && confidence >= 80) {
      blacklist.push({
        name: c.company_name,
        reason: result.reasoning,
        company_type: companyType,
        added: new Date().toISOString()
      });
      blacklisted++;
    }

    classified.push({ ...c, ...result, domain });

    // Human-like delay between detail pages (prevents rate limiting)
    await new Promise(r => setTimeout(r, 3000 + Math.random() * 4000));
  }

  await context.close();
  await browser.close();

  // ── Save ──
  db.last_scrape = new Date().toISOString();
  db.last_updated = new Date().toISOString();
  saveJSON(DB_PATH, db);
  saveJSON(BLACKLIST_PATH, blacklist);

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

  // ── Send to Webhook ──
  if (WEBHOOK_URL) {
    const azList = db.companies.filter(c => c.in_arizona && c.st_user);
    
    const payload = JSON.stringify({
      date: endDate,
      jobs_scraped: jobs.length,
      companies_checked: companies.length,
      new_st_users: newST,
      new_az_st_users: newAZ,
      blacklisted_today: blacklisted,
      db_total: db.companies.length,
      db_az_st_users: azList.length,
      new_confirmed: db.companies
        .filter(c => c.first_seen === endDate && c.st_user)
        .map(c => ({
          company: c.company_name,
          domain: c.domain,
          location: c.location,
          job: c.first_job,
          confidence: c.confidence,
          insight: c.reasoning
        })),
      az_confirmed: db.companies
        .filter(c => c.in_arizona && c.st_user)
        .map(c => ({
          company: c.company_name,
          domain: c.domain,
          location: c.location,
          job: c.first_job,
          date: (c.reviewed_at || c.first_seen || '').split('T')[0]
        }))
    });

    try {
      const webhookReq = https.request(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        timeout: 15000
      }, (res) => {
        console.log(`Webhook sent: ${res.statusCode}`);
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
}

main().catch(err => {
  console.error('FATAL:', err.message);
  process.exit(1);
});
