const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const inputDir = "Input";
const outputDir = "Output";

const MAX_SHEET_NAME_LENGTH = 31;

// --- Helpers ---

function cleanSheetName(name) {
  let clean = name.replace(/\.json$/i, "");
  clean = clean.replace(/[:\\/?*\[\]]/g, "_");
  if (clean.length > MAX_SHEET_NAME_LENGTH) {
    clean = clean.substring(0, MAX_SHEET_NAME_LENGTH);
  }
  return clean;
}

// Parse "HH:MM:SS" or "HHHH:MM:SS" to seconds
function parseDuration(durationStr) {
  if (!durationStr || typeof durationStr !== "string") return 0;
  // Handle cases like "2966:35:53" or "01:33:11.493717"
  // Remove potential milliseconds for simple summing or keep them?
  // Usually H:M:S is good enough.
  const parts = durationStr.split(":");
  if (parts.length < 2) return 0;

  const hours = parseInt(parts[0], 10) || 0;
  const minutes = parseInt(parts[1], 10) || 0;
  const seconds = parseFloat(parts[2]) || 0;

  return hours * 3600 + minutes * 60 + seconds;
}

// Format seconds to "HH:MM:SS" (floor seconds)
function formatDuration(totalSeconds) {
  const hours = Math.floor(totalSeconds / 3600);
  const remainder = totalSeconds % 3600;
  const minutes = Math.floor(remainder / 60);
  const seconds = Math.floor(remainder % 60);

  return `${hours}:${minutes.toString().padStart(2, "0")}:${seconds
    .toString()
    .padStart(2, "0")}`;
}

// --- Aggregation Logic ---

function aggregateData(filename, data) {
  if (!Array.isArray(data) || data.length === 0) return [];

  // Identify file type to apply specific rules
  const name = filename.replace(/\.json$/i, "").toUpperCase();
  console.log(`Aggregating for: ${name}`);

  // Group by 'Nombre_Unidad'
  const groups = {};
  for (const row of data) {
    // Use 'Nombre_Unidad' or fallback to 'General' if not present (though prompt says by unit)
    // Some files might use 'Departamento' instead of Unit if it's the department report?
    // User asked: "un acamulado del anio por unidad".
    // Checks columns.
    const unit = row["Nombre_Unidad"] || "N/A";
    if (!groups[unit]) {
      groups[unit] = { rows: [], count: 0 };
    }
    groups[unit].rows.push(row);
    groups[unit].count++;
  }

  const aggregatedRows = [];

  for (const [unit, groupData] of Object.entries(groups)) {
    if (unit === "N/A") continue; // Skip if no unit found (or handle differently)

    const summaryRow = {
      Mes_Anio: "Acumulado 2025",
      Nombre_Unidad: unit,
    };

    // Rule-based aggregation
    // STRICTER ORDERING: Check specific suffixes first!
    if (
      name.includes("MANTENIMIENTO") ||
      name.includes("TECNOLOGIA") ||
      name.includes("LOST_AND_FOUND")
    ) {
      processGenericCounts(summaryRow, groupData.rows);
    } else if (name.includes("GLITCHES")) {
      processGlitches(summaryRow, groupData.rows);
    } else if (
      name === "TICKETS_GENERAL" ||
      name === "DATOS_GENERALES_UNIDAD_DEPARTAMENTO" ||
      name.startsWith("DATOS_GENERALES")
    ) {
      // Only for the specific time-tracking files
      processTicketsGeneral(summaryRow, groupData.rows);
    } else if (name === "ETIQUETAS_UNIDAD_DEPARTAMENTO") {
      // This one has 'tiempo_promedio_productivo' but might structure differently.
      // Let's use generic or specific if needed. Generic might try to sum averages which is wrong.
      // For now, let's treat as generic but be careful.
      processGenericCounts(summaryRow, groupData.rows);
    } else {
      // Fallback: try to sum obvious columns
      processGenericCounts(summaryRow, groupData.rows);
    }

    aggregatedRows.push(summaryRow);
  }

  return aggregatedRows;
}

function processGenericCounts(summaryRow, rows) {
  // Sum all columns starting with 'Cantidad' or 'Total'
  // Copy arbitrary keys from first row just to see structure
  const keys = Object.keys(rows[0]);

  for (const key of keys) {
    if (key === "Mes_Anio" || key === "Nombre_Unidad") continue;

    // Check if value is numeric-ish
    const isNumeric = rows.every((r) => !isNaN(Number(r[key])));

    if (isNumeric) {
      const sum = rows.reduce((acc, r) => acc + (Number(r[key]) || 0), 0);
      summaryRow[key] = sum;
    } else {
      summaryRow[key] = ""; // Leave text columns empty
    }
  }
}

function processGlitches(summaryRow, rows) {
  // Similar to generic but ensure specific fields
  processGenericCounts(summaryRow, rows);
}

function processTicketsGeneral(summaryRow, rows) {
  // Cantidad_Tickets: Sum
  summaryRow["Cantidad_Tickets"] = rows.reduce(
    (acc, r) => acc + (Number(r["Cantidad_Tickets"]) || 0),
    0
  );

  // Total_Tiempo_Productivo: Time Sum
  const totalTimeProdSec = rows.reduce((acc, r) => {
    const val = r["Total_Tiempo_Productivo"];
    const sec = parseDuration(val);
    // console.log(`    Time val: ${val} -> ${sec}s`);
    return acc + sec;
  }, 0);

  summaryRow["Total_Tiempo_Productivo"] = formatDuration(totalTimeProdSec);

  // Promedio_Tiempo_Productivo: Calc (Total Prod Time / Count Tickets)
  if (summaryRow["Cantidad_Tickets"] > 0) {
    const avgProdSec = totalTimeProdSec / summaryRow["Cantidad_Tickets"];
    summaryRow["Promedio_Tiempo_Productivo"] = formatDuration(avgProdSec);
  } else {
    summaryRow["Promedio_Tiempo_Productivo"] = "0:00:00";
  }

  // Promedio_Tiempo_Estimado: Weighted Average
  // We don't have Total Estimated. We have Avg Estimated.
  // Total Estimated ~= Sum(Avg_Est * Ticket_Count)
  let totalEstSec = 0;
  for (const r of rows) {
    const estSec = parseDuration(r["Promedio_Tiempo_Estimado"]);
    const count = Number(r["Cantidad_Tickets"]) || 0;
    totalEstSec += estSec * count;
  }

  if (summaryRow["Cantidad_Tickets"] > 0) {
    const avgEstSec = totalEstSec / summaryRow["Cantidad_Tickets"];
    summaryRow["Promedio_Tiempo_Estimado"] = formatDuration(avgEstSec);

    // Porcentaje_Cumplimiento: (Avg Est / Avg Prod) * 100
    // Or (Total Est / Total Prod) * 100
    if (totalTimeProdSec > 0) {
      const pct = (totalEstSec / totalTimeProdSec) * 100;
      summaryRow["Porcentaje_Cumplimiento"] = pct.toFixed(2);
    } else {
      summaryRow["Porcentaje_Cumplimiento"] = "0.00";
    }
  } else {
    summaryRow["Promedio_Tiempo_Estimado"] = "0:00:00";
    summaryRow["Porcentaje_Cumplimiento"] = "0.00";
  }
}

// --- Data Filling Logic ---

const MONTHS_2025 = [
  "Enero 2025",
  "Febrero 2025",
  "Marzo 2025",
  "Abril 2025",
  "Mayo 2025",
  "Junio 2025",
  "Julio 2025",
  "Agosto 2025",
  "Septiembre 2025",
  "Octubre 2025",
  "Noviembre 2025",
  "Diciembre 2025",
];

function fillMissingMonths(data) {
  if (!Array.isArray(data) || data.length === 0) return data;

  // Check if data has 'Mes_Anio' and 'Nombre_Unidad'
  if (
    !data[0].hasOwnProperty("Mes_Anio") ||
    !data[0].hasOwnProperty("Nombre_Unidad")
  ) {
    return data;
  }

  // Get all unique units
  const units = [...new Set(data.map((r) => r.Nombre_Unidad))];
  const filledData = [];

  // Template for zero values based on first row
  const keys = Object.keys(data[0]);

  for (const unit of units) {
    // Filter rows for this unit
    const unitRows = data.filter((r) => r.Nombre_Unidad === unit);
    const unitMonths = new Set(unitRows.map((r) => r.Mes_Anio));

    for (const month of MONTHS_2025) {
      if (unitMonths.has(month)) {
        // Add existing row(s) -- simplified assumption: one row per month per unit
        // If duplicates (unlikely for these queries), we take them.
        const existing = unitRows.find((r) => r.Mes_Anio === month);
        if (existing) filledData.push(existing);
      } else {
        // Create zero row
        const zeroRow = {};
        for (const key of keys) {
          if (key === "Mes_Anio") zeroRow[key] = month;
          else if (key === "Nombre_Unidad") zeroRow[key] = unit;
          else if (
            key.startsWith("Cantidad") ||
            key.startsWith("Total_Tickets") ||
            key === "insertId" ||
            key === "affectedRows"
          )
            zeroRow[key] = 0;
          else if (key.includes("Tiempo"))
            zeroRow[key] =
              "00:00:00"; // formatting matters? '0:00:00' or '00:00:00'. Script uses '0:00:00' usually.
          else if (key.includes("Porcentaje")) zeroRow[key] = "0.00";
          else zeroRow[key] = 0; // default numeric
        }
        filledData.push(zeroRow);
      }
    }
  }

  // Sort logic?
  // The current loop structure pushes in Month order (Enero -> Dic) for each Unit.
  // But we might want to group by Unit?
  // The loop `for (const unit of units)` effectively groups by Unit.
  // Inside, it goes `for (const month of MONTHS_2025)`, so it sorts by Month.
  // Result: Unit A (Jan-Dec), Unit B (Jan-Dec).
  // The original SQL was ORDER BY Year, Month, Unit. -> Jan (Unit A, Unit B), Feb (Unit A, Unit B).
  // User screenshot shows grouped by Month? No, screenshot shows "Febrero... Palacio", "Marzo... Palacio".
  // Screenshot 1767725266085 shows "Enero 2025 Pierre", "Febrero 2025 Palacio", "Febrero 2025 Princess".
  // Wait, the original SQL order is Month, then Unit.
  // My `fillMissingMonths` changes it to Unit, then Month.

  // Let's stick to Unit -> Month for now as it's cleaner for "Missing months per unit".
  // OR we can flatten and sort by Month index if needed.
  // Let's keep Unit -> Month?
  // Actually, looking at the user request "en los generales puedes colocar los meses aunque sean 0",
  // usually readable reports are either by Month (comparing units) or by Unit (chronological).
  // The SQL output was mixed:
  /*
        ORDER BY
        YEAR(mt.fCreacionTicket),
        MONTH(mt.fCreacionTicket),
        Nombre_Unidad;
    */
  // That means: Jan Unit 1, Jan Unit 2... Feb Unit 1, Feb Unit 2.
  // If I change to Unit 1 (Jan-Dec), Unit 2 (Jan-Dec), that changes the layout.
  // Let's try to restore the Month-first sort order.

  const monthIndex = {};
  MONTHS_2025.forEach((m, i) => (monthIndex[m] = i));

  filledData.sort((a, b) => {
    const ma = monthIndex[a.Mes_Anio];
    const mb = monthIndex[b.Mes_Anio];
    if (ma !== mb) return ma - mb;
    if (a.Nombre_Unidad < b.Nombre_Unidad) return -1;
    if (a.Nombre_Unidad > b.Nombre_Unidad) return 1;
    return 0;
  });

  return filledData;
}

function processDirectoryToWorkbook(directoryPath, outputFilename) {
  if (!fs.existsSync(directoryPath)) return;

  const items = fs.readdirSync(directoryPath, { withFileTypes: true });
  const jsonFiles = items.filter(
    (item) => item.isFile() && item.name.toLowerCase().endsWith(".json")
  );

  if (jsonFiles.length === 0) return;

  const workbook = XLSX.utils.book_new();
  let hasSheets = false;

  console.log(`Creating ${outputFilename} with ${jsonFiles.length} files...`);

  for (const file of jsonFiles) {
    const filePath = path.join(directoryPath, file.name);
    try {
      const content = fs.readFileSync(filePath, "utf8");
      let data = JSON.parse(content);

      if (Array.isArray(data) && data.length > 0) {
        const sheetName = cleanSheetName(file.name);

        // 0. Fill Missing Months (Only for General files likely? User said "En los generales")
        // Check if it has the required columns to be safe
        if (data[0].Mes_Anio && data[0].Nombre_Unidad) {
          data = fillMissingMonths(data);
        }

        // 1. Generate Aggregated Rows
        const accumulated = aggregateData(file.name, data);

        // 2. Combine Data with spacer
        const finalData = [...data];

        if (accumulated.length > 0) {
          finalData.push({}); // Spacer
          finalData.push({}); // Spacer
          finalData.push(...accumulated);
        }

        const worksheet = XLSX.utils.json_to_sheet(finalData);

        let finalSheetName = sheetName;
        let counter = 1;
        while (workbook.SheetNames.includes(finalSheetName)) {
          finalSheetName = `${sheetName.substring(
            0,
            MAX_SHEET_NAME_LENGTH - 3
          )}_${counter}`;
          counter++;
        }

        XLSX.utils.book_append_sheet(workbook, worksheet, finalSheetName);
        hasSheets = true;
        console.log(`  + Added sheet: ${finalSheetName}`);
      }
    } catch (err) {
      console.error(`  ! Error processing ${file.name}:`, err.message);
    }
  }

  if (hasSheets) {
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
    XLSX.writeFile(workbook, path.join(outputDir, outputFilename));
    console.log(`Saved workbook: ${outputFilename}`);
  }
}

function main() {
  if (!fs.existsSync(inputDir)) {
    console.error(`Input directory '${inputDir}' does not exist.`);
    return;
  }

  // 1. Process Root JSON files -> General.xlsx
  console.log("Processing General files...");
  processDirectoryToWorkbook(inputDir, "General.xlsx");

  // 2. Process Unit Subdirectories -> {Unit}.xlsx
  const items = fs.readdirSync(inputDir, { withFileTypes: true });
  const subdirs = items.filter((item) => item.isDirectory());

  console.log(`Found ${subdirs.length} unit directories.`);

  for (const dir of subdirs) {
    console.log(`Processing Unit: ${dir.name}`);
    const unitPath = path.join(inputDir, dir.name);
    // For Unit specific files, they might already be just for that unit,
    // so accumulation might be just one row per table.
    // User asked for "where each file.json of that unit is a sheet".
    // The logic holds.
    const excelName = `${dir.name}.xlsx`;
    processDirectoryToWorkbook(unitPath, excelName);
  }

  console.log("Excel conversion complete.");
}

main();
