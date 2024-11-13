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
    dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
    visitDate: isNaN(record["Visit Date"]) ? record["Visit Date"].trim() : formatDate(record["Visit Date"]),
    charges: [record["Charge1"], record["Charge2"], record["Charge3"]].filter(Boolean),
    pmdProvider: record["Provider"] ? record["Provider"].split(",")[0].trim().toLowerCase() : null,
    midlevelProvider: record["Midlevel Visit: Visit done in coordination with midlevel:"]
        ? record["Midlevel Visit: Visit done in coordination with midlevel:"].toLowerCase().trim()
        : null
}));

// Modify the logic to use midlevelProvider if it's available
const pmdProviderToMatch = (pmdRecord) => {
    // Use midlevelProvider if it exists, otherwise fallback to pmdProvider
    return pmdRecord.midlevelProvider || pmdRecord.pmdProvider;
};

// Normalize and format records from ECW
const ecwRecords = ecwData.map(record => {
    const lastName = record["Rendering Provider"] ? record["Rendering Provider"].split(",")[0].trim() : null;
    const firstName = record["Resource Provider"] ? record["Resource Provider"].split(",")[1]?.trim() : null;
    const ecwProviders = lastName && firstName ? `${firstName} ${lastName}`.toLowerCase() : null;

    return {
        name: record["Patient"].toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Start Date of Service"]) ? record["Start Date of Service"].trim() : formatDate(record["Start Date of Service"]),
        cpt: record["CPT Code"],
        claimNo: record["Claim No"],
        ecwProviders
    };
});

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

        let claimNo = null;
        const missingCPTs = [];
        let ecwProviders = null;
        let providerMismatch = false;

        // Loop through each charge in PMD and check for matches in ECW
        pmdRecord.charges.forEach(charge => {
            const match = ecwMatches.find(ecwRecord => ecwRecord.cpt === charge);
            if (match) {
                claimNo = match.claimNo;
                ecwProviders = match.ecwProviders; // Set ecwProviders based on the match
            } else {
                missingCPTs.push(charge);
            }
        });

        // Check if providers match (case-insensitive); if not, flag provider mismatch
        const providerToCompare = pmdProviderToMatch(pmdRecord);
        if (
            providerToCompare &&
            ecwProviders &&
            providerToCompare !== ecwProviders
        ) {
            providerMismatch = true;
        }

        // Handle combined status for missing CPTs and provider mismatch
        if (missingCPTs.length > 0 || providerMismatch) {
            mistakesCount++;
            let statusMessage = "";

            // Add details to the status message based on issues found
            if (missingCPTs.length > 0) {
                statusMessage += `missing CPTs: ${missingCPTs.join(", ")}`;
            }
            if (providerMismatch) {
                statusMessage += (statusMessage ? " and " : "") + "providers not matched";
            }

            mistakeRecords.push({
                ...pmdRecord,
                status: statusMessage,
                missingCPTs,
                claimNo,
                pmdProvider: providerToCompare, // Use the matched provider (either midlevel or main provider)
                ecwProviders
            });
        } else {
            // If all matches
            matchedRecords.push({
                ...pmdRecord,
                status: "matched",
                charges: pmdRecord.charges,
                claimNo,
                pmdProvider: providerToCompare,
                ecwProviders
            });
        }
    } else {
        missingCount++;
        missingRecords.push({
            ...pmdRecord,
            status: "missing",
            charges: pmdRecord.charges,
            pmdProvider: pmdRecord.pmdProvider
        });
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
