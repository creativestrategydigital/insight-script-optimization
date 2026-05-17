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
  // DAILY CAPACITY SETUP (40 visits/day max = 4 auditors × 10 visits)
  // ======================================================

  const DAILY_MAX_CAPACITY =
    NUM_AUDITORS * MAX_VISITS_PER_AUDITOR;  // 4 * 10 = 40

  const dailyTargets = {};
  const auditsByDate = {};

  for (const d of AUDITOR_DATES) {
    dailyTargets[d] = DAILY_MAX_CAPACITY;
    auditsByDate[d] = [];
  }

  // Check eligibility per date and adjust targets
  const eligibleByDate = {};
  for (const d of AUDITOR_DATES) {
    eligibleByDate[d] = 0;
  }
  for (const v of filteredVisits) {
    for (const ad of v.eligibleAuditDates) {
      if (eligibleByDate[ad] !== undefined) {
        eligibleByDate[ad]++;
      }
    }
  }

  // Adjust targets: if a day has fewer eligible visits than capacity, reduce
  for (const d of AUDITOR_DATES) {
    if (eligibleByDate[d] === 0) {
      dailyTargets[d] = 0;
    } else if (eligibleByDate[d] < DAILY_MAX_CAPACITY) {
      dailyTargets[d] = eligibleByDate[d];
    }
  }

  console.log('\n======================');
  console.log('DAILY CAPACITY');
  console.log('======================');
  for (const d of AUDITOR_DATES) {
    const eligible = eligibleByDate[d];
    const cap = dailyTargets[d];
    if (eligible === 0) {
      console.log(`⚠ ${d}: ELIGIBLE=0 → CAPACITY=0`);
    } else if (cap < DAILY_MAX_CAPACITY) {
      console.log(`⚠ ${d}: ELIGIBLE=${eligible} < MAX=${DAILY_MAX_CAPACITY} → CAPACITY=${cap}`);
    } else {
      console.log(`✓ ${d}: ELIGIBLE=${eligible}, CAPACITY=${cap}`);
    }
  }

  // ======================================================
  // SELECT VISITS WITH DATE-AWARE EQUITABLE DISTRIBUTION
  // ======================================================

  const selectedMain = [];
  const selectedBuffer = [];

  // Group filtered visits by SR, Tier, AND eligible audit date
  const visitsBySRTierDate = {};
  for (const v of filteredVisits) {
    const tier = v['Verif tiering'] || 'Unknown';
    for (const ad of v.eligibleAuditDates) {
      const groupKey = `${v.SR}|${tier}|${ad}`;
      if (!visitsBySRTierDate[groupKey]) {
        visitsBySRTierDate[groupKey] = [];
      }
      visitsBySRTierDate[groupKey].push(v);
    }
  }

  // Track how many visits are assigned per day
  const assignedPerDay = {};
  for (const d of AUDITOR_DATES) {
    assignedPerDay[d] = 0;
  }

  // Track selected visit IDs to avoid duplicates
  const selectedIds = new Set();

  console.log('\n======================');
  console.log('EQUITABLE SELECTION BY SR AND TIER (DATE-AWARE)');
  console.log('======================');

  for (const tier of tiers) {
    console.log(`\n--- ${tier} ---`);

    for (const sr of finalSRList) {
      const key = `${sr}|${tier}`;
      const target = targetsBySRandTier[key] || 0;
      if (target === 0) continue;

      // Collect all available visits for this SR/Tier, grouped by date
      const visitsByDate = {};
      for (const d of AUDITOR_DATES) {
        const groupKey = `${sr}|${tier}|${d}`;
        const pool = (visitsBySRTierDate[groupKey] || []).filter(
          v => !selectedIds.has(v['DB-ID'])
        );
        if (pool.length > 0) {
          visitsByDate[d] = shuffle(pool);
        }
      }

      let selected = 0;

      // Round-robin across dates sorted by remaining capacity (least filled first)
      while (selected < target) {
        // Get dates with remaining capacity and available visits
        const availableDates = AUDITOR_DATES.filter(
          d => visitsByDate[d] && visitsByDate[d].length > 0 &&
               assignedPerDay[d] < dailyTargets[d]
        );

        if (availableDates.length === 0) break;

        // Sort by remaining capacity descending (fill least-full days first)
        availableDates.sort((a, b) =>
          (dailyTargets[b] - assignedPerDay[b]) - (dailyTargets[a] - assignedPerDay[a])
        );

        // Pick one visit from the day with most remaining capacity
        const bestDate = availableDates[0];
        const v = visitsByDate[bestDate].pop();

        if (selectedIds.has(v['DB-ID'])) continue;

        selectedIds.add(v['DB-ID']);
        v.AuditDate = bestDate;
        auditsByDate[bestDate].push(v);
        assignedPerDay[bestDate]++;
        selectedMain.push(v);
        selected++;
      }

      const availableTotal = (filteredBySRandTier[key] || []).length;
      console.log(
        `${sr} | ${tier}: Target=${target}, Available=${availableTotal}, Selected=${selected}`
      );
    }
  }

  // If under target, try to fill from remaining pool
  if (selectedMain.length < TARGET_AUDITS) {
    console.log(`\n⚠ Under target: ${selectedMain.length}/${TARGET_AUDITS}. Filling from remaining pool...`);

    const remainingPool = shuffle(
      filteredVisits.filter(v => !selectedIds.has(v['DB-ID']))
    );

    for (const v of remainingPool) {
      if (selectedMain.length >= TARGET_AUDITS) break;

      // Find a date with capacity
      const possibleDates = v.eligibleAuditDates.filter(
        d => assignedPerDay[d] < dailyTargets[d]
      );

      if (possibleDates.length === 0) continue;

      // Pick the date with most remaining capacity
      possibleDates.sort((a, b) =>
        (dailyTargets[b] - assignedPerDay[b]) - (dailyTargets[a] - assignedPerDay[a])
      );

      const bestDate = possibleDates[0];
      selectedIds.add(v['DB-ID']);
      v.AuditDate = bestDate;
      auditsByDate[bestDate].push(v);
      assignedPerDay[bestDate]++;
      selectedMain.push(v);
    }
  }

  console.log(`\n✓ Total selected for main sample: ${selectedMain.length}`);

  console.log('\n======================');
  console.log('ASSIGNMENT SUMMARY');
  console.log('======================');

  let totalInMain = 0;
  for (const d of AUDITOR_DATES) {
    const count = auditsByDate[d].length;
    const cap = dailyTargets[d];
    totalInMain += count;

    if (count === 0 && cap === 0) {
      console.log(`⚠ ${d}: No audits (no eligible visits)`);
    } else if (count === cap) {
      console.log(`✓ ${d}: ${count}/${cap} (full)`);
    } else {
      console.log(`⚠ ${d}: ${count}/${cap}`);
    }
  }

  console.log(`\n✓ Total audits in main file: ${totalInMain}`);

  // Fill buffer with non-selected visits from the 20 selected SRs
  const bufferPool = filteredVisits.filter(v => !selectedIds.has(v['DB-ID']));
  const shuffledBufferPool = shuffle(bufferPool);
  selectedBuffer.push(...shuffledBufferPool.slice(0, BUFFER_SIZE));

  console.log(`✓ Buffer: ${selectedBuffer.length} visits (max ${BUFFER_SIZE})`);

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
// node generate_audit_routes.js "Abidjan_Mai_16_28.xls" > audit_plan.txt