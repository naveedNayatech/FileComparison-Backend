const express = require("express");
const router = express.Router();
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const moment = require("moment");

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});
const upload = multer({ storage: storage });

// Helper function to format Excel serial dates
const formatDate = (serial) => {
    const excelStartDate = new Date(1899, 11, 30);
    const date = new Date(excelStartDate.getTime() + (serial * 86400000));
    return moment(date).format("MM/DD/YYYY");
};

// Function to parse and compare Excel files
function compareExcelFiles(pmdFileBuffer, ecwFileBuffer) {
    const pmdWorkbook = xlsx.read(pmdFileBuffer, { type: "buffer" });
    const ecwWorkbook = xlsx.read(ecwFileBuffer, { type: "buffer" });
    
    const pmdSheet = pmdWorkbook.Sheets[pmdWorkbook.SheetNames[0]];
    const ecwSheet = ecwWorkbook.Sheets[ecwWorkbook.SheetNames[0]];

    const pmdData = xlsx.utils.sheet_to_json(pmdSheet);
    const ecwData = xlsx.utils.sheet_to_json(ecwSheet);

    // Normalize and format records from PMD
    const pmdRecords = pmdData.map(record => ({
        visitId: record["Visit ID"],
        name: `${record["Patient Last"]}, ${record["Patient First"]}`.toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"] : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Visit Date"]) ? record["Visit Date"] : formatDate(record["Visit Date"]),
        charges: [record["Charge1"], record["Charge2"], record["Charge3"]].filter(Boolean), // Collect charges
    }));

    // Normalize and format records from ECW
    const ecwRecords = ecwData.map(record => ({
        name: record["Patient"].toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"] : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Start Date of Service"]) ? record["Start Date of Service"] : formatDate(record["Start Date of Service"]),
        cpt: record["CPT Code"] // Assuming ECW has a "CPT Code" field for each row
    }));

    // Initialize stats and record arrays
    const matchedRecords = [];
    const missingRecords = [];
    const mistakeRecords = [];
    let matchedCount = 0;
    let missingCount = 0;
    let mistakesCount = 0;

    pmdRecords.forEach(pmdRecord => {
        // Find matching records in ECW based on name, dob, and visit date
        const ecwMatches = ecwRecords.filter(
            ecwRecord => 
                ecwRecord.name === pmdRecord.name &&
                ecwRecord.dob === pmdRecord.dob &&
                ecwRecord.visitDate === pmdRecord.visitDate
        );

        if (ecwMatches.length > 0) {
            matchedCount++;

            // Collect CPT codes found and missing CPTs
            const foundCPTs = [];
            const missingCPTs = [];

            pmdRecord.charges.forEach(charge => {
                const match = ecwMatches.find(ecwRecord => ecwRecord.cpt === charge);
                if (match) {
                    foundCPTs.push(match.cpt);
                } else {
                    missingCPTs.push(charge);
                }
            });

            if (missingCPTs.length === 0) {
                matchedRecords.push({
                    ...pmdRecord,
                    status: "matched",
                    charges: pmdRecord.charges,
                    foundCPTs // CPTs that were successfully matched
                });
            } else {
                mistakesCount++;
                mistakeRecords.push({
                    ...pmdRecord,
                    status: `missing CPTs: ${missingCPTs.join(", ")}`,
                    missingCPTs,
                    foundCPTs // CPTs that were matched partially
                });
            }
        } else {
            missingCount++;
            missingRecords.push({ ...pmdRecord, status: "missing", charges: pmdRecord.charges });
        }
    });

    // Prepare comparison results
    return {
        stats: {
            matchedCount,
            missingCount,
            mistakesCount
        },
        matchedRecords,
        missingRecords,
        mistakeRecords
    };
}

// Define the route for file comparison
router.post("/compare", upload.fields([{ name: "pmdFile" }, { name: "ecwFile" }]), (req, res) => {
    const pmdFile = req.files['pmdFile'] ? req.files['pmdFile'][0] : null;
    const ecwFile = req.files['ecwFile'] ? req.files['ecwFile'][0] : null;

    if (!pmdFile || !ecwFile) {
        return res.status(400).json({ error: 'Both files are required' });
    }

    const pmdFileBuffer = fs.readFileSync(pmdFile.path);
    const ecwFileBuffer = fs.readFileSync(ecwFile.path);

    try {
        const results = compareExcelFiles(pmdFileBuffer, ecwFileBuffer);

        res.json({
            message: "Hospital Files Comparison complete",
            stats: results.stats,
            matchedRecords: results.matchedRecords,
            missingRecords: results.missingRecords,
            mistakeRecords: results.mistakeRecords
        });
    } catch (error) {
        res.status(500).send("Error comparing Excel files: " + error.message);
    } finally {
        // Clean up uploaded files
        fs.unlinkSync(pmdFile.path);
        fs.unlinkSync(ecwFile.path);
    }
});

module.exports = router;
