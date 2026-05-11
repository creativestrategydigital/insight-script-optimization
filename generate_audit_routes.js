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
const BUFFER_SIZE = 120;
 
// TOTAL SAMPLE
const TOTAL_SAMPLE =
  TARGET_AUDITS +
  BUFFER_SIZE;
 
// ======================================================
// DATE FUNCTIONS
// ======================================================
 
function parseDate(d) {
 
  if (d instanceof Date) {
    return d;
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
  // LOAD CSV
  // ======================================================
 
  const raw =
    fs.readFileSync(
      filename,
      'utf8'
    );
 
  const rows =
    parseCSV(raw);
 
  console.log(
    `✓ Loaded ${rows.length} rows`
  );
 
  // ======================================================
  // PREPARE VISITS
  // ======================================================
 
  const visits = [];
 
  for (const row of rows) {
 
    const lat =
      parseFloat(
        row.Latitude
          ?.replace(',', '.')
      );
 
    const lon =
      parseFloat(
        row.Longitude
          ?.replace(',', '.')
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

    for (const offset of [2, 3]) {

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
  // GROUP BY SR
  // ======================================================
 
  const bySR = {};
 
  for (const v of visits) {
 
    if (!bySR[v.SR]) {
      bySR[v.SR] = [];
    }
 
    bySR[v.SR].push(v);
  }
 
  // ======================================================
  // PROPORTIONAL SAMPLE
  // ======================================================
 
  const selectedMain = [];
 
  const selectedBuffer = [];
 
  const totalUniverse =
    visits.length;
 
  console.log(
    '\n======================'
  );
 
  console.log(
    'SAMPLE PLAN'
  );
 
  console.log(
    '======================'
  );
 
  for (
    const [sr, srVisits]
    of Object.entries(bySR)
  ) {
 
    const proportion =
      srVisits.length /
      totalUniverse;
 
    const totalSampleSize =
      Math.max(
        1,
        Math.round(
          proportion *
          TOTAL_SAMPLE
        )
      );
 
    const mainSampleSize =
      Math.round(
        totalSampleSize *
        TARGET_AUDITS /
        TOTAL_SAMPLE
      );
 
    const bufferSampleSize =
      totalSampleSize -
      mainSampleSize;
 
    const randomized =
      shuffle(srVisits);
 
    const mainPart =
      randomized.slice(
        0,
        mainSampleSize
      );
 
    const bufferPart =
      randomized.slice(
        mainSampleSize,
        mainSampleSize +
        bufferSampleSize
      );
 
    selectedMain.push(
      ...mainPart
    );
 
    selectedBuffer.push(
      ...bufferPart
    );
 
    console.log(
      `${sr} | Universe=${srVisits.length} | Main=${mainPart.length} | Buffer=${bufferPart.length}`
    );
  }
 
  console.log(
    `\n✓ Main sample: ${selectedMain.length}`
  );
 
  console.log(
    `✓ Buffer sample: ${selectedBuffer.length}`
  );
 
  console.log(
    `✓ Total sample: ${selectedMain.length + selectedBuffer.length}`
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

  for (const v of selectedMain) {

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

  console.log(
    'ELIGIBLE VS TARGETS'
  );

  console.log(
    '======================'
  );

  for (const d of AUDITOR_DATES) {

    const eligible =
      eligibleByDate[d];

    const target =
      dailyTargets[d];

    const actualTarget =
      Math.min(target, eligible);

    if (actualTarget < target) {

      console.log(
        `⚠ ${d}: TARGET=${target} but ELIGIBLE=${eligible} → ADJUSTED TO ${actualTarget}`
      );

      dailyTargets[d] =
        actualTarget;
    } else {

      console.log(
        `✓ ${d}: TARGET=${target}, ELIGIBLE=${eligible}`
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