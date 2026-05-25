#!/usr/bin/env node
/**
 * LinkedIn Company Deduplication
 * Reads raw scrape results, filters new companies vs persistent database.
 * Outputs a batch of NEW companies ready for AI classification.
 *
 * Usage:
 *   node scripts/dedup-linkedin.js
 *     --input data/linkedin_raw.json
 *     --db    data/linkedin_companies.json
 *     --output data/linkedin_pending.json
 */

const fs = require('fs');

const INPUT = process.argv.includes('--input')
  ? process.argv[process.argv.indexOf('--input') + 1]
  : 'data/linkedin_raw.json';
const DB = process.argv.includes('--db')
  ? process.argv[process.argv.indexOf('--db') + 1]
  : 'data/linkedin_companies.json';
const OUTPUT = process.argv.includes('--output')
  ? process.argv[process.argv.indexOf('--output') + 1]
  : 'data/linkedin_pending.json';

function normalize(str) {
  return (str || '')
    .toLowerCase()
    .replace(/[,.\-'&]/g, '')
    .replace(/\b(llc|inc|incorporat|corp|corporation|ltd|limited|co|company|dba|lp|group|partners|services|holdings)\b/gi, '')
    .replace(/\s+/g, '')
    .trim();
}

function loadDB() {
  if (fs.existsSync(DB)) {
    try { return JSON.parse(fs.readFileSync(DB, 'utf-8')); }
    catch { return { companies: [], last_updated: null }; }
  }
  return { companies: [], last_updated: null };
}

function saveDB(db) {
  fs.writeFileSync(DB, JSON.stringify(db, null, 2));
}

function main() {
  if (!fs.existsSync(INPUT)) {
    console.error(`INPUT not found: ${INPUT}`);
    console.log(JSON.stringify({ status: 'error', message: `Input file not found: ${INPUT}` }));
    process.exit(1);
  }

  const raw = JSON.parse(fs.readFileSync(INPUT, 'utf-8'));
  const db = loadDB();
  
  console.error(`Raw companies: ${raw.companies.length}`);
  console.error(`DB has: ${db.companies.length} reviewed companies`);
  
  // Normalize existing company names for lookup
  const existingNames = new Set();
  for (const c of db.companies) {
    existingNames.add(normalize(c.company_name));
  }
  
  // Filter to new companies only
  const newCompanies = raw.companies.filter(c => {
    const key = normalize(c.company_name);
    if (!key || key.length < 3) return false; // Skip empty/too-short names
    if (existingNames.has(key)) return false; // Already reviewed
    return true;
  });
  
  console.error(`New companies to classify: ${newCompanies.length}`);
  console.error(`Skipped: ${raw.companies.length - newCompanies.length} (already in DB)`);
  
  // For each new company, check if the location contains Arizona
  const azCities = ['phoenix', 'tucson', 'mesa', 'scottsdale', 'gilbert', 'chandler', 
                     'glendale', 'tempe', 'peoria', 'surprise', 'yuma', 'flagstaff',
                     'arizona', 'az'];
  
  const pending = newCompanies.map(c => {
    const locLower = (c.location || '').toLowerCase();
    const inArizona = azCities.some(city => locLower.includes(city));
    return { ...c, in_arizona: inArizona };
  });
  
  // Sort: Arizona companies first
  pending.sort((a, b) => (b.in_arizona ? 1 : 0) - (a.in_arizona ? 1 : 0));
  
  // Write pending batch
  const output = {
    generated_at: new Date().toISOString(),
    total_pending: pending.length,
    in_arizona: pending.filter(c => c.in_arizona).length,
    companies: pending
  };
  
  fs.writeFileSync(OUTPUT, JSON.stringify(output, null, 2));
  
  console.error(`Arizona companies: ${output.in_arizona}`);
  console.error(`Saved to ${OUTPUT}`);
  
  // Update DB's last scrape time
  db.last_scrape = new Date().toISOString();
  saveDB(db);
  
  console.log(JSON.stringify({
    status: 'completed',
    total_pending: pending.length,
    in_arizona: output.in_arizona,
    output_file: OUTPUT
  }));
}

main();
