#!/usr/bin/env node
/**
 * Merge AI-reviewed results back into the persistent database
 *
 * Usage:
 *   node scripts/merge-linkedin.js
 *     --input data/linkedin_reviewed.json
 *     --db    data/linkedin_companies.json
 */

const fs = require('fs');

const INPUT = process.argv.includes('--input')
  ? process.argv[process.argv.indexOf('--input') + 1]
  : 'data/linkedin_reviewed.json';
const DB = process.argv.includes('--db')
  ? process.argv[process.argv.indexOf('--db') + 1]
  : 'data/linkedin_companies.json';

function main() {
  if (!fs.existsSync(INPUT)) {
    console.error(`INPUT not found: ${INPUT}`);
    console.log(JSON.stringify({ status: 'error', message: `Input file not found: ${INPUT}` }));
    process.exit(1);
  }

  const reviewed = JSON.parse(fs.readFileSync(INPUT, 'utf-8'));
  let db = { companies: [], last_updated: null };
  if (fs.existsSync(DB)) {
    try { db = JSON.parse(fs.readFileSync(DB, 'utf-8')); }
    catch { /* use default */ }
  }

  let added = 0;
  for (const company of reviewed.companies) {
    // Only add if classified (not skipped)
    if (!company.company_type) continue;
    
    // Check for duplicates by normalized name
    const norm = (company.company_name || '').toLowerCase().trim()
      .replace(/[,.\-'&]/g, '')
      .replace(/\b(llc|inc|corp|ltd|co)\b/gi, '')
      .replace(/\s+/g, '');
    
    const exists = db.companies.some(c => {
      const cn = (c.company_name || '').toLowerCase().trim()
        .replace(/[,.\-'&]/g, '')
        .replace(/\b(llc|inc|corp|ltd|co)\b/gi, '')
        .replace(/\s+/g, '');
      return cn === norm;
    });
    
    if (!exists) {
      db.companies.push({
        ...company,
        status: 'reviewed',
        reviewed_at: new Date().toISOString()
      });
      added++;
    }
  }

  db.last_updated = new Date().toISOString();
  fs.writeFileSync(DB, JSON.stringify(db, null, 2));

  // Stats
  const az = db.companies.filter(c => c.in_arizona && c.st_user && c.company_type === 'home_service').length;
  const totalST = db.companies.filter(c => c.st_user && c.company_type === 'home_service').length;

  console.error(`Added: ${added} | Skipped (duplicates): ${reviewed.companies.length - added}`);
  console.error(`DB total: ${db.companies.length}`);
  console.error(`AZ home service ST users: ${az} | All ST users: ${totalST}`);

  console.log(JSON.stringify({ status: 'completed', added, db_total: db.companies.length, az_st_users: az }));
}

main();
