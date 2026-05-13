const fs = require('fs');
const XLSX = require('xlsx');
 
// ======================================================
// CONFIGURATION
// ======================================================
 
const AUDITOR_DATES = [
  "18/05/2026",
  "19/05/2026",
  "20/05/2026",
  "21/05/2026",
  "22/05/2026",
  "23/05/2026",
  "25/05/2026",
  "26/05/2026",
  "27/05/2026",
  "28/05/2026",
  "29/05/2026",
  "30/05/2026"
];
 
const AUDITOR_DATE_SET =
  new Set(AUDITOR_DATES);
 
// ======================================================
// AUDIT PARAMETERS
// ======================================================
 
const NUM_AUDITORS = 4;
 
const MIN_VISITS_PER_AUDITOR = 8;
 
const MAX_VISITS_PER_AUDITOR = 10;
 
// MAIN SAMPLE
const TARGET_AUDITS = 400;

// BUFFER
const BUFFER_PERCENT = 0.3;
const BUFFER_SIZE = 120;

// ======================================================
// DATE FUNCTIONS
// ======================================================
 
function parseDate(d) {

  if (d instanceof Date) {
    return d;
  }

  // EXCEL SERIAL DATE (e.g., 46158)
  if (
    typeof d === 'number' ||
    /^\d+$/.test(d)
  ) {

    const serial =
      Number(d);

    // Excel serial date: days since 1900-01-01
    // But Excel thinks 1900 was a leap year (bug), so we adjust
    const epoch =
      new Date(1899, 11, 30);

    const date =
      new Date(
        epoch.getTime() +
        serial *
          24 *
          60 *
          60 *
          1000
      );

    return date;
  }

  const [dd, mm, yyyy] =
    d.toString()
      .trim()
      .split('/')
      .map(Number);

  return new Date(
    yyyy,
    mm - 1,
    dd
  );
}
 
function formatDate(d) {
 
  return `${String(
    d.getDate()
  ).padStart(2, '0')}/${String(
    d.getMonth() + 1
  ).padStart(2, '0')}/${d.getFullYear()}`;
}
 
// ======================================================
// RANDOMIZER
// ======================================================
 
function shuffle(array) {
 
  const arr = [...array];
 
  for (
    let i = arr.length - 1;
    i > 0;
    i--
  ) {
 
    const j = Math.floor(
      Math.random() *
      (i + 1)
    );
 
    [arr[i], arr[j]] =
      [arr[j], arr[i]];
  }
 
  return arr;
}
 
// ======================================================
// DISTANCE
// ======================================================
 
function haversine(
  lat1,
  lon1,
  lat2,
  lon2
) {
 
  const R = 6371;
 
  const dLat =
    (lat2 - lat1) *
    Math.PI / 180;
 
  const dLon =
    (lon2 - lon1) *
    Math.PI / 180;
 
  const a =
    Math.sin(dLat / 2) *
    Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) *
    Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) *
    Math.sin(dLon / 2);
 
  return (
    R *
    2 *
    Math.atan2(
      Math.sqrt(a),
      Math.sqrt(1 - a)
    )
  );
}
 
// ======================================================
// ROUTE OPTIMIZATION
// ======================================================
 
function optimizeRoute(visits) {
 
  if (visits.length <= 1) {
 
    return {
      ordered: visits,
      distance: 0
    };
  }
 
  const ordered = [visits[0]];
 
  const remaining =
    visits.slice(1);
 
  while (remaining.length > 0) {
 
    const last =
      ordered[
        ordered.length - 1
      ];
 
    let nearestIdx = 0;
 
    let minDist = Infinity;
 
    for (
      let i = 0;
      i < remaining.length;
      i++
    ) {
 
      const dist =
        haversine(
          last.Latitude,
          last.Longitude,
          remaining[i].Latitude,
          remaining[i].Longitude
        );
 
      if (dist < minDist) {
 
        minDist = dist;
 
        nearestIdx = i;
      }
    }
 
    ordered.push(
      remaining.splice(
        nearestIdx,
        1
      )[0]
    );
  }
 
  let totalDistance = 0;
 
  for (
    let i = 0;
    i < ordered.length - 1;
    i++
  ) {
 
    totalDistance +=
      haversine(
        ordered[i].Latitude,
        ordered[i].Longitude,
        ordered[i + 1].Latitude,
        ordered[i + 1].Longitude
      );
  }
 
  return {
    ordered,
    distance: totalDistance
  };
}
 
// ======================================================
// CSV PARSER
// ======================================================
 
function parseCSV(content) {
 
  const lines =
    content.split(
      /\r\n|\n|\r/
    );
 
  const headerLine =
    lines[0];
 
  const separator =
    headerLine.includes(';')
      ? ';'
      : ',';
 
  const headers =
    headerLine
      .split(separator)
      .map(h => h.trim());
 
  const rows = [];
 
  for (
    let i = 1;
    i < lines.length;
    i++
  ) {
 
    const line =
      lines[i].trim();
 
    if (!line) continue;
 
    const values =
      line.split(separator);
 
    const row = {};
 
    headers.forEach(
      (h, idx) => {
 
        row[h] =
          values[idx]
            ?.trim() || '';
      }
    );
 
    rows.push(row);
  }
 
  return rows;
}
 
// ======================================================
// MAIN
// ======================================================
 
function main() {
 
  const filename =
    process.argv[2];
 
  if (!filename) {
 
    console.log(
      'Usage: node audit_sampling.js file.csv'
    );
 
    process.exit(1);
  }
 
  // ======================================================
  // LOAD FILE (CSV OR EXCEL)
  // ======================================================

  let rows;

  if (
    filename.endsWith('.csv')
  ) {

    const raw =
      fs.readFileSync(
        filename,
        'utf8'
      );

    rows = parseCSV(raw);
  } else if (
    filename.endsWith('.xls') ||
    filename.endsWith('.xlsx')
  ) {

    const workbook =
      XLSX.readFile(filename);

    const firstSheet =
      workbook.SheetNames[0];

    const worksheet =
      workbook.Sheets[firstSheet];

    rows =
      XLSX.utils.sheet_to_json(
        worksheet
      );
  } else {

    console.log(
      'Error: File must be .csv, .xls, or .xlsx'
    );

    process.exit(1);
  }

  console.log(
    `✓ Loaded ${rows.length} rows`
  );

  // DEBUG: Check first rows
  console.log(
    '\n=== DEBUG FIRST 3 ROWS ==='
  );

  for (
    let i = 0;
    i < Math.min(3, rows.length);
    i++
  ) {

    console.log(
      `Row ${i}: DATE="${rows[i].DATE}" (type: ${typeof rows[i].DATE}), Lat="${rows[i].Latitude}", Lon="${rows[i].Longitude}"`
    );

    const parsedDate =
      parseDate(rows[i].DATE);

    console.log(
      `  Parsed date: ${parsedDate}`
    );
  }

  console.log(
    '========================\n'
  );

  // ======================================================
  // PREPARE VISITS
  // ======================================================
 
  const visits = [];

  for (const row of rows) {

    const lat =
      parseFloat(
        String(row.Latitude)
          .replace(',', '.')
      );

    const lon =
      parseFloat(
        String(row.Longitude)
          .replace(',', '.')
      );

    if (
      isNaN(lat) ||
      isNaN(lon)
    ) {
      continue;
    }
 
    const visitDate =
      parseDate(row.DATE);

    // D+2 OR D+3 CONSTRAINT
    const eligibleAuditDates =
      [];

    for (const offset of [2]) {

      const auditDate =
        new Date(visitDate);

      auditDate.setDate(
        auditDate.getDate() +
        offset
      );

      const auditDateStr =
        formatDate(auditDate);

      if (
        AUDITOR_DATE_SET.has(
          auditDateStr
        )
      ) {

        eligibleAuditDates.push(
          auditDateStr
        );
      }
    }

    // SKIP IF NO ELIGIBLE DATES
    if (
      eligibleAuditDates.length === 0
    ) {
      continue;
    }

    row.Latitude = lat;

    row.Longitude = lon;

    row.eligibleAuditDates =
      eligibleAuditDates;

    visits.push(row);
  }
 
  console.log(
    `✓ Eligible visits: ${visits.length}`
  );
 
  // ======================================================
  // GROUP BY SR AND TIER
  // ======================================================

  const bySR = {};
  const bySRandTier = {};

  for (const v of visits) {

    if (!bySR[v.SR]) {
      bySR[v.SR] = [];
    }
    bySR[v.SR].push(v);

    const tier = v['Verif tiering'] || 'Unknown';
    const key = `${v.SR}|${tier}`;

    if (!bySRandTier[key]) {
      bySRandTier[key] = [];
    }
    bySRandTier[key].push(v);
  }

  // ======================================================
  // CALCULATE TARGETS PER SR AND TIER (EQUITABLE DISTRIBUTION)
  // ======================================================

  const srList = Object.keys(bySR).sort((a, b) => a.localeCompare(b));
  const tiers = ['Tier 1', 'Tier 2', 'Tier 3'];

  // Step 1: Calculate total visits per SR
  const totalVisitsBySR = {};
  for (const sr of srList) {
    totalVisitsBySR[sr] = bySR[sr].length;
  }

  // Step 2: Select TOP 20 SR with most visits
  const TOP_SR_COUNT = 20;
  const sortedSR = [...srList].sort((a, b) => totalVisitsBySR[b] - totalVisitsBySR[a]);
  const selectedSRList = sortedSR.slice(0, TOP_SR_COUNT);

  console.log(
    '\n======================'
  );
  console.log(
    `TOP ${TOP_SR_COUNT} SR SELECTED (by total visits)`
  );
  console.log(
    '======================'
  );
  for (let i = 0; i < sortedSR.length; i++) {
    const sr = sortedSR[i];
    const marker = i < TOP_SR_COUNT ? '✓' : '✗';
    console.log(`${marker} ${sr}: ${totalVisitsBySR[sr]} visits`);
  }

  // Step 3: Filter visits to only include selected SRs
  const filteredVisits = visits.filter(v => selectedSRList.includes(v.SR));
  const filteredBySR = {};
  const filteredBySRandTier = {};

  for (const v of filteredVisits) {
    if (!filteredBySR[v.SR]) {
      filteredBySR[v.SR] = [];
    }
    filteredBySR[v.SR].push(v);

    const tier = v['Verif tiering'] || 'Unknown';
    const key = `${v.SR}|${tier}`;
    if (!filteredBySRandTier[key]) {
      filteredBySRandTier[key] = [];
    }
    filteredBySRandTier[key].push(v);
  }

  // Recalculate with filtered data
  const finalSRList = selectedSRList;
  let globalTotalVisits = filteredVisits.length;

  // Step 4: Calculate total target for each SR (proportional to their visits)
  const targetBySR = {};
  for (const sr of finalSRList) {
    const proportion = filteredBySR[sr].length / globalTotalVisits;
    targetBySR[sr] = Math.round(proportion * TARGET_AUDITS);
  }

  // Step 5: Distribute each SR's target EQUALLY between their 3 Tiers
  const targetsBySRandTier = {};

  for (const sr of finalSRList) {
    const srTarget = targetBySR[sr];

    // Get available visits per tier for this SR
    const availableByTier = {};
    for (const tier of tiers) {
      const key = `${sr}|${tier}`;
      availableByTier[tier] = filteredBySRandTier[key] ? filteredBySRandTier[key].length : 0;
    }

    const totalAvailable = availableByTier['Tier 1'] + availableByTier['Tier 2'] + availableByTier['Tier 3'];

    if (totalAvailable === 0) {
      continue;
    }

    // Start with equal distribution (as much as possible)
    let remainingTarget = Math.min(srTarget, totalAvailable);
    const tierTargets = {};

    // First pass: try to give equal amount to each tier with available visits
    const tiersWithVisits = tiers.filter(t => availableByTier[t] > 0);
    const basePerTier = Math.floor(remainingTarget / tiersWithVisits.length);

    for (const tier of tiers) {
      if (availableByTier[tier] > 0) {
        tierTargets[tier] = Math.min(basePerTier, availableByTier[tier]);
        remainingTarget -= tierTargets[tier];
      } else {
        tierTargets[tier] = 0;
      }
    }

    // Second pass: distribute remaining to tiers that can take more
    for (const tier of tiers) {
      if (remainingTarget <= 0) break;
      const canAdd = availableByTier[tier] - tierTargets[tier];
      if (canAdd > 0) {
        const add = Math.min(canAdd, remainingTarget);
        tierTargets[tier] += add;
        remainingTarget -= add;
      }
    }

    // Store targets
    for (const tier of tiers) {
      const key = `${sr}|${tier}`;
      targetsBySRandTier[key] = tierTargets[tier] || 0;
    }
  }

  // Step 6: Adjust to ensure we hit exactly TARGET_AUDITS (400)
  let calculatedTotal = 0;
  for (const key of Object.keys(targetsBySRandTier)) {
    calculatedTotal += targetsBySRandTier[key];
  }

  let diff = TARGET_AUDITS - calculatedTotal;

  if (diff !== 0) {
    // Sort all SR-Tier combinations by available (descending) for adjustment
    const allCombinations = [];
    for (const sr of finalSRList) {
      for (const tier of tiers) {
        const key = `${sr}|${tier}`;
        const available = filteredBySRandTier[key] ? filteredBySRandTier[key].length : 0;
        const target = targetsBySRandTier[key] || 0;
        if (available > 0) {
          allCombinations.push({ key, sr, tier, available, target });
        }
      }
    }

    allCombinations.sort((a, b) => b.available - a.available);

    // Add or remove to match TARGET_AUDITS
    for (const combo of allCombinations) {
      if (diff === 0) break;
      if (diff > 0 && combo.target < combo.available) {
        const add = Math.min(diff, combo.available - combo.target);
        targetsBySRandTier[combo.key] = combo.target + add;
        diff -= add;
      } else if (diff < 0 && combo.target > 0) {
        const remove = Math.min(Math.abs(diff), combo.target);
        targetsBySRandTier[combo.key] = combo.target - remove;
        diff += remove;
      }
    }
  }

  // ======================================================
  // SELECT VISITS BASED ON TARGETS
  // ======================================================

  const selectedMain = [];
  const selectedBuffer = [];

  console.log(
    '\n======================'
  );

  console.log(
    'EQUITABLE SELECTION BY SR AND TIER'
  );

  console.log(
    '======================'
  );

  for (const tier of tiers) {
    console.log(`\n--- ${tier} ---`);

    for (const sr of finalSRList) {
      const key = `${sr}|${tier}`;
      const available = filteredBySRandTier[key] || [];
      const target = targetsBySRandTier[key] || 0;

      if (available.length > 0) {
        // Shuffle and select
        const shuffled = shuffle(available);
        const selected = shuffled.slice(0, Math.min(target, available.length));

        selectedMain.push(...selected);

        console.log(
          `${sr} | ${tier}: Target=${target}, Available=${available.length}, Selected=${selected.length}`
        );
      }
    }
  }

  // Add remaining visits from filtered data if we haven't reached TARGET_AUDITS exactly
  const currentTotal = selectedMain.length;
  if (currentTotal < TARGET_AUDITS) {
    console.log(`\n⚠ Under target: ${currentTotal}/${TARGET_AUDITS}. Adding from remaining pool...`);

    // Find remaining eligible visits not yet selected (from filtered pool)
    const selectedSet = new Set(selectedMain.map(v => v['DB-ID']));
    const remaining = filteredVisits.filter(v => !selectedSet.has(v['DB-ID']));
    const shuffledRemaining = shuffle(remaining);
    const needed = TARGET_AUDITS - currentTotal;

    selectedMain.push(...shuffledRemaining.slice(0, needed));
  }

  console.log(
    `\n✓ Total selected for main sample: ${selectedMain.length}`
  );
 
  // ======================================================
  // FIXED DAILY TARGETS
  // ======================================================
 
  const dailyTargets = {};
 
  const BASE_TARGET =
    Math.floor(
      TARGET_AUDITS /
      AUDITOR_DATES.length
    );
 
  let remaining =
    TARGET_AUDITS -
    (
      BASE_TARGET *
      AUDITOR_DATES.length
    );
 
  for (const d of AUDITOR_DATES) {
 
    dailyTargets[d] =
      BASE_TARGET;
  }
 
  // DISTRIBUTE REMAINING
  for (const d of AUDITOR_DATES) {
 
    if (remaining <= 0) {
      break;
    }
 
    dailyTargets[d]++;
 
    remaining--;
  }
 
  console.log(
    '\n======================'
  );
 
  console.log(
    'FIXED DAILY TARGETS'
  );
 
  console.log(
    '======================'
  );
 
  for (const d of AUDITOR_DATES) {
 
    console.log(
      `${d} : ${dailyTargets[d]} audits`
    );
  }
 
  // ======================================================
  // CHECK ELIGIBILITY VS TARGETS
  // ======================================================

  const eligibleByDate = {};

  for (const d of AUDITOR_DATES) {

    eligibleByDate[d] = 0;
  }

  for (const v of visits) {

    for (
      const ad of
      v.eligibleAuditDates
    ) {

      if (
        eligibleByDate[ad] !==
        undefined
      ) {

        eligibleByDate[ad]++;
      }
    }
  }

  console.log(
    '\n======================'
  );

  const DAILY_MAX_CAPACITY =
    NUM_AUDITORS *
    MAX_VISITS_PER_AUDITOR;  // 4 * 10 = 40

  // FIRST PASS: Set target to min(eligible, capacity) or 0 if no eligible
  for (const d of AUDITOR_DATES) {

    const eligible =
      eligibleByDate[d];

    if (eligible === 0) {

      dailyTargets[d] = 0;

      console.log(
        `⚠ ${d}: ELIGIBLE=0 → ADJUSTED TO 0`
      );
    } else if (
      eligible >=
      DAILY_MAX_CAPACITY
    ) {

      dailyTargets[d] =
        DAILY_MAX_CAPACITY;

      console.log(
        `✓ ${d}: ELIGIBLE=${eligible}, CAPACITY=${DAILY_MAX_CAPACITY}`
      );
    } else {

      dailyTargets[d] =
        eligible;

      console.log(
        `⚠ ${d}: ELIGIBLE=${eligible} < CAPACITY=${DAILY_MAX_CAPACITY} → ADJUSTED TO ${eligible}`
      );
    }
  }

  // ======================================================
  // ASSIGN AUDITS TO DAYS
  // ======================================================

  const auditsByDate = {};

  for (const d of AUDITOR_DATES) {

    auditsByDate[d] = [];
  }
 
  const randomizedMain =
    shuffle(selectedMain);
 
  for (const v of randomizedMain) {
 
    let assigned = false;
 
    const possibleDates =
      shuffle(
        v.eligibleAuditDates
      );
 
    for (
      const ad of possibleDates
    ) {
 
      if (
        auditsByDate[ad]
          .length <
        dailyTargets[ad]
      ) {
 
        v.AuditDate = ad;
 
        auditsByDate[ad]
          .push(v);
 
        assigned = true;
 
        break;
      }
    }
 
    // FALLBACK
    if (!assigned) {
 
      const fallback =
        possibleDates[0];
 
      v.AuditDate =
        fallback;
 
      auditsByDate[
        fallback
      ].push(v);
    }
  }

  // ======================================================
  // ASSIGNMENT SUMMARY
  // ======================================================

  console.log(
    '\n======================'
  );

  console.log(
    'ASSIGNMENT SUMMARY'
  );

  console.log(
    '======================'
  );

  for (const d of AUDITOR_DATES) {

    const count =
      auditsByDate[d].length;

    const target =
      dailyTargets[d];

    console.log(
      `${d}: ASSIGNED=${count}, TARGET=${target}`
    );
  }

  // ======================================================
  // KEEP ALL ASSIGNED AUDITS IN MAIN FILE
  // ======================================================

  console.log(
    '\n======================'
  );

  console.log(
    'KEEPING ALL ASSIGNED AUDITS IN MAIN FILE'
  );

  console.log(
    '======================'
  );

  // Calculate how many audits we have per day
  for (const d of AUDITOR_DATES) {
    const count = auditsByDate[d].length;
    const target = dailyTargets[d];

    if (count > target) {
      console.log(`⚠ ${d}: ${count} audits assigned (exceeds target ${target})`);
    } else if (count < target && count > 0) {
      console.log(`⚠ ${d}: ${count} audits assigned (below target ${target})`);
    } else if (count === target && count > 0) {
      console.log(`✓ ${d}: ${count} audits assigned (exact target)`);
    } else {
      console.log(`⚠ ${d}: No audits assigned`);
    }
  }

  // Calculate total audits in main file
  let totalInMain = 0;
  for (const d of AUDITOR_DATES) {
    totalInMain += auditsByDate[d].length;
  }

  console.log(`\n✓ Total audits in main file: ${totalInMain}`);

  // Buffer is created from the pool of non-selected visits if needed
  // For now, we keep the buffer empty or fill it with remaining eligible visits
  let bufferCountTotal = 0;

  if (selectedBuffer.length === 0 && bufferCountTotal < BUFFER_SIZE) {
    // If buffer is empty, we can optionally fill it from remaining visits
    // But for equitable distribution, we keep buffer separate
    console.log(`✓ Buffer: ${selectedBuffer.length} visits`);
  }

  // ======================================================
  // BUILD FINAL OUTPUT
  // ======================================================

  const finalRows = [];

  for (
    const auditDate of
    AUDITOR_DATES
  ) {

    const dayVisits =
      auditsByDate[
        auditDate
      ];

    if (
      dayVisits.length === 0
    ) {
      continue;
    }

 
    // SORT GEO
    const sorted =
      [...dayVisits].sort(
        (a, b) =>
          a.Latitude -
          b.Latitude
      );

    // ======================================================
    // SPLIT BY AUDITOR (EVEN DISTRIBUTION)
    // ======================================================

    const baseSize =
      Math.floor(
        sorted.length /
        NUM_AUDITORS
      );

    const remainder =
      sorted.length %
      NUM_AUDITORS;

    let currentIndex = 0;

    for (
      let auditor = 1;
      auditor <=
      NUM_AUDITORS;
      auditor++
    ) {

      const chunkSize =
        baseSize +
        (auditor <= remainder ? 1 : 0);

      const chunk =
        sorted.slice(
          currentIndex,
          currentIndex +
          chunkSize
        );

      currentIndex +=
        chunkSize;

      if (
        chunk.length === 0
      ) {
        continue;
      }

      // SECURITY CHECK
      if (
        chunk.length >
        MAX_VISITS_PER_AUDITOR
      ) {

        console.log(
          `⚠ WARNING: Auditor ${auditor} has ${chunk.length} visits on ${auditDate}`
        );
      }

      if (
        chunk.length <
        MIN_VISITS_PER_AUDITOR
      ) {

        console.log(
          `⚠ WARNING: Auditor ${auditor} has only ${chunk.length} visits on ${auditDate}`
        );
      }

      // ROUTE OPTIMIZATION
      const optimized =
        optimizeRoute(
          chunk
        );
 
      optimized.ordered
        .forEach(
          (v, idx) => {
 
            finalRows.push({
 
              Auditor:
                auditor,
 
              AuditDate:
                auditDate,
 
              Sequence:
                idx + 1,
 
              OriginalVisitDate:
                v.DATE,
 
              SR:
                v.SR,

              Tier:
                v['Verif tiering'],

              Territory:
                v.Territory,
 
              Outlet:
                v['Outlet Name'],
 
              SEM_ID:
                v['SEM ID'],
 
              DB_ID:
                v['DB-ID'],
 
              Region:
                v.Region,
 
              Channel:
                v['New Channel'],
 
              Telephone:
                v.Telephone,
 
              Latitude:
                v.Latitude,
 
              Longitude:
                v.Longitude
            });
          }
        );
    }
  }
 
  // ======================================================
  // EXPORT MAIN FILE
  // ======================================================
 
  const mainWorkbook =
    XLSX.utils.book_new();
 
  const mainSheet =
    XLSX.utils.json_to_sheet(
      finalRows
    );
 
  XLSX.utils.book_append_sheet(
    mainWorkbook,
    mainSheet,
    'Audit Routes'
  );
 
  XLSX.writeFile(
    mainWorkbook,
    'Audit_Main_400.xlsx'
  );
 
  // ======================================================
  // EXPORT BUFFER FILE
  // ======================================================
 
  const bufferRows =
    selectedBuffer.map(
      v => ({
 
        AuditDate:
          v.AuditDate,

        OriginalVisitDate:
          v.DATE,
 
        SR:
          v.SR,
 
        Territory:
          v.Territory,
 
        Outlet:
          v['Outlet Name'],
 
        SEM_ID:
          v['SEM ID'],
 
        DB_ID:
          v['DB-ID'],
 
        Region:
          v.Region,
 
        Channel:
          v['New Channel'],
 
        Telephone:
          v.Telephone,
 
        Latitude:
          v.Latitude,
 
        Longitude:
          v.Longitude
      })
    );
 
  const bufferWorkbook =
    XLSX.utils.book_new();
 
  const bufferSheet =
    XLSX.utils.json_to_sheet(
      bufferRows
    );
 
  XLSX.utils.book_append_sheet(
    bufferWorkbook,
    bufferSheet,
    'Buffer'
  );
 
  XLSX.writeFile(
    bufferWorkbook,
    'Audit_Buffer_120.xlsx'
  );
 
  // ======================================================
  // FINAL SUMMARY
  // ======================================================
 
  console.log(
    '\n======================'
  );
 
  console.log(
    'FINAL RESULTS'
  );
 
  console.log(
    '======================'
  );
 
  console.log(
    `✓ Main audits exported: ${finalRows.length}`
  );
 
  console.log(
    `✓ Buffer exported: ${bufferRows.length}`
  );
 
  console.log(
    `✓ Total selected: ${finalRows.length + bufferRows.length}`
  );
 
  console.log(
    `✓ Average audits/day: ${(
      finalRows.length /
      AUDITOR_DATES.length
    ).toFixed(2)}`
  );
 
  console.log(
    `✓ Average audits/auditor/day: ${(
      finalRows.length /
      AUDITOR_DATES.length /
      NUM_AUDITORS
    ).toFixed(2)}`
  );
 
  console.log(
    '\n✓ Main file: Audit_Main_400.xlsx'
  );
 
  console.log(
    '✓ Buffer file: Audit_Buffer_120.xlsx'
  );
}
 
main();
// node generate_audit_routes.js "Abidjan_Mai_16_28.csv" > audit_plan.txt