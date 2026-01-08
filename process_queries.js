require("dotenv").config();
const mysql = require("mysql2/promise");
const fs = require("fs");
const path = require("path");

const inputFile = "Consultas.txt";
const outputDir = "Input";

async function main() {
  try {
    console.log("Starting query processing...");

    // 1. Read the input file
    const fileContent = fs.readFileSync(inputFile, "utf-8");
    const lines = fileContent.split(/\r?\n/);

    const sections = [];
    let currentTitle = null;
    let currentSql = [];

    for (const line of lines) {
      if (line.trim().endsWith("--------------------------")) {
        // If we have a previous section, save it
        if (currentTitle) {
          sections.push({ title: currentTitle, sql: currentSql.join("\n") });
        }
        // Start new section
        currentTitle = line.replace(/--------------------------$/, "").trim();
        currentSql = [];
      } else {
        if (currentTitle) {
          currentSql.push(line);
        }
      }
    }
    // Push the last section
    if (currentTitle && currentSql.length > 0) {
      sections.push({ title: currentTitle, sql: currentSql.join("\n") });
    }

    console.log(`Found ${sections.length} sections to process.`);

    // 2. Connect to database
    const connection = await mysql.createConnection({
      host: process.env.DB_HOST,
      user: process.env.DB_USERNAME,
      password: process.env.DB_PASSWORD,
      database: process.env.DB_DATABASE,
      port: process.env.DB_PORT,
      multipleStatements: true,
    });

    console.log("Connected to database.");

    // Fetch Units 1, 2, 3
    const [units] = await connection.query(
      "SELECT idUnidad, nombreUnidad FROM unidad WHERE idUnidad IN (1, 2, 3)"
    );
    const unitMap = {};
    for (const u of units) {
      unitMap[u.idUnidad] = u.nombreUnidad;
    }
    console.log("Units found:", unitMap);

    // 3. Process each section
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    for (const section of sections) {
      console.log(`Processing: ${section.title}`);
      const isUnitSpecific = section.sql.includes("u.idUnidad = 1");

      if (isUnitSpecific) {
        console.log(
          ` -> Detected unit specific query. Iterating for units 1-3...`
        );
        for (const unitId of [1, 2, 3]) {
          const unitName = unitMap[unitId] || `Unidad_${unitId}`;
          const modifiedSql = section.sql.replace(
            /u\.idUnidad\s*=\s*1/g,
            `u.idUnidad = ${unitId}`
          );

          try {
            const [results] = await connection.query(modifiedSql);
            let dataToExport = results;

            if (Array.isArray(results)) {
              const rowSets = results.filter((r) => Array.isArray(r));
              if (rowSets.length > 0) {
                dataToExport = rowSets[rowSets.length - 1];
              }
            }

            // Create unit folder
            const unitFolder = path.join(
              outputDir,
              unitName.replace(/[^a-zA-Z0-9]/g, "_")
            );
            if (!fs.existsSync(unitFolder)) {
              fs.mkdirSync(unitFolder);
            }

            const safeTitle = section.title.replace(/[^a-zA-Z0-9]/g, "_");
            const filename = `${safeTitle}.json`;
            const outputPath = path.join(unitFolder, filename);

            fs.writeFileSync(outputPath, JSON.stringify(dataToExport, null, 2));
            console.log(`    -> Saved to ${outputPath}`);
          } catch (err) {
            console.error(`    -> Error for Unit ${unitId}:`, err.message);
          }
        }
      } else {
        // General query
        try {
          const [results] = await connection.query(section.sql);
          let dataToExport = results;

          if (Array.isArray(results)) {
            const rowSets = results.filter((r) => Array.isArray(r));
            if (rowSets.length > 0) {
              dataToExport = rowSets[rowSets.length - 1];
            }
          }

          const safeTitle = section.title.replace(/[^a-zA-Z0-9]/g, "_");
          const filename = `${safeTitle}.json`;
          const outputPath = path.join(outputDir, filename);

          fs.writeFileSync(outputPath, JSON.stringify(dataToExport, null, 2));
          console.log(` -> Saved to ${outputPath}`);
        } catch (err) {
          console.error(
            `Error processing section '${section.title}':`,
            err.message
          );
        }
      }
    }

    await connection.end();
    console.log("Done.");
  } catch (err) {
    console.error("Fatal error:", err);
  }
}

main();
