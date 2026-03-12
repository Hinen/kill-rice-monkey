import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';

const TOTAL = 5000;
const CONCURRENCY = 5;
const OUTPUT_DIR = path.join(process.cwd(), 'captcha-samples', 'melon');
const PROD_ID = '212444';
const SCHEDULE_NO = '100001';

const COOKIE = process.argv[2];
if (!COOKIE) {
  console.error('Usage: node captcha-collect.mjs "<cookie>"');
  process.exit(1);
}

fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const existing = fs.readdirSync(OUTPUT_DIR).filter(f => f.endsWith('.png'));
let startIdx = existing.length;
console.log(`Starting from index ${startIdx}, target: ${TOTAL}`);

let collected = startIdx;
let errors = 0;

function fetchCaptcha(idx) {
  return new Promise((resolve, reject) => {
    const t = Date.now() + idx;
    const url = `/reservation/ajax/captChaImage.json?prodId=${PROD_ID}&scheduleNo=${SCHEDULE_NO}&t=${t}`;
    const options = {
      hostname: 'ticket.melon.com',
      port: 443,
      path: url,
      method: 'GET',
      headers: {
        'Cookie': COOKIE,
        'Referer': `https://ticket.melon.com/reservation/popup/onestop.htm`,
        'X-Requested-With': 'XMLHttpRequest',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.CAPTIMAGE) {
            const filename = `captcha_${String(idx).padStart(5, '0')}.png`;
            const filepath = path.join(OUTPUT_DIR, filename);
            fs.writeFileSync(filepath, Buffer.from(json.CAPTIMAGE, 'base64'));
            resolve({ idx, filename, size: json.CAPTIMAGE.length });
          } else {
            reject(new Error(`No CAPTIMAGE in response for idx ${idx}: ${Object.keys(json).join(',')}`));
          }
        } catch (e) {
          reject(new Error(`Parse error for idx ${idx}: ${e.message}`));
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(10000, () => {
      req.destroy();
      reject(new Error(`Timeout for idx ${idx}`));
    });
    req.end();
  });
}

async function collectBatch(startFrom, batchSize) {
  const promises = [];
  for (let i = 0; i < batchSize; i++) {
    promises.push(
      fetchCaptcha(startFrom + i).catch(err => {
        errors++;
        return { error: err.message };
      })
    );
  }
  return Promise.all(promises);
}

async function main() {
  console.log(`Collecting ${TOTAL - startIdx} CAPTCHA samples...`);
  const startTime = Date.now();

  while (collected < TOTAL) {
    const remaining = TOTAL - collected;
    const batchSize = Math.min(CONCURRENCY, remaining);
    const results = await collectBatch(collected, batchSize);

    const successes = results.filter(r => !r.error);
    collected += successes.length;

    if (collected % 100 === 0 || collected >= TOTAL) {
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      const rate = (successes.length > 0) ? ((collected - startIdx) / ((Date.now() - startTime) / 1000)).toFixed(1) : '0';
      console.log(`[${elapsed}s] Collected: ${collected}/${TOTAL} (errors: ${errors}, rate: ${rate}/s)`);
    }

    if (errors > 50) {
      console.error('Too many errors, stopping.');
      break;
    }

    await new Promise(r => setTimeout(r, 100));
  }

  console.log(`Done. Total collected: ${collected}, errors: ${errors}`);
}

main().catch(console.error);
