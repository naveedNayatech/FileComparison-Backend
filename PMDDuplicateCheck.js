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

// Compare function to check if PMD and ECW records match
const compareRecords = (pmdRecord, ecwRecord) => {
    return pmdRecord.name === ecwRecord.name &&
           pmdRecord.dob === ecwRecord.dob &&
           pmdRecord.visitDate === ecwRecord.visitDate;
};

// Function to compare Excel files and return duplicates
function compareExcelFilesForDuplicates(pmdFileBuffer, ecwFileBuffer) {
    const pmdWorkbook = xlsx.read(pmdFileBuffer, { type: "buffer" });
    const ecwWorkbook = xlsx.read(ecwFileBuffer, { type: "buffer" });

    const pmdSheet = pmdWorkbook.Sheets[pmdWorkbook.SheetNames[0]];
    const ecwSheet = ecwWorkbook.Sheets[ecwWorkbook.SheetNames[0]];

    const pmdData = xlsx.utils.sheet_to_json(pmdSheet);
    const ecwData = xlsx.utils.sheet_to_json(ecwSheet);

    const pmdRecords = pmdData.map(record => ({
        visitId: record["Visit ID"],
        name: `${record["Patient Last"]}, ${record["Patient First"]}`.toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Visit Date"]) ? record["Visit Date"].trim() : formatDate(record["Visit Date"]),
        charges: [record["Charge1"], record["Charge2"], record["Charge3"]].filter(Boolean)
    }));

    const ecwRecords = ecwData.map(record => ({
        name: record["Patient"].toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Start Date of Service"]) ? record["Start Date of Service"].trim() : formatDate(record["Start Date of Service"]),
        cpt: record["CPT Code"],
        claimNo: record["Claim No"]
    }));

    const results = [];

    // Loop through each PMD record
    pmdRecords.forEach(pmdRecord => {
        pmdRecord.charges.forEach(charge => {
            const matchedECWRecords = ecwRecords.filter(ecwRecord =>
                compareRecords(pmdRecord, ecwRecord) && ecwRecord.cpt === charge
            );

            const status = matchedECWRecords.length > 0 ? "matched" : "not matched";
            const timesPresentInECW = matchedECWRecords.length === 1 ? 1 : (matchedECWRecords.length === 2 ? 2 : matchedECWRecords.length);

            results.push({
                visitId: pmdRecord.visitId,
                name: pmdRecord.name,
                dob: pmdRecord.dob,
                visitDate: pmdRecord.visitDate,
                charge: charge,
                status: status,
                timesPresentInECW: timesPresentInECW
            });
        });
    });

    return { results };
}

router.post("/find-duplicates", upload.fields([{ name: "pmdFile" }, { name: "ecwFile" }]), (req, res) => {
    const pmdFile = req.files['pmdFile'] ? req.files['pmdFile'][0] : null;
    const ecwFile = req.files['ecwFile'] ? req.files['ecwFile'][0] : null;

    if (!pmdFile || !ecwFile) {
        return res.status(400).json({ error: 'Both files are required' });
    }

    const pmdFileBuffer = fs.readFileSync(pmdFile.path);
    const ecwFileBuffer = fs.readFileSync(ecwFile.path);

    try {
        const results = compareExcelFilesForDuplicates(pmdFileBuffer, ecwFileBuffer);

        // Filter results to only include records where timesPresentInECW is 2 or more
        const filteredResults = results.results.filter(record => record.timesPresentInECW >= 2);

        res.json({
            message: "Duplicate records detection complete",
            stats: { total: filteredResults.length },
            records: filteredResults
        });
    } catch (error) {
        res.status(500).send("Error finding duplicate records: " + error.message);
    } finally {
        fs.unlinkSync(pmdFile.path);
        fs.unlinkSync(ecwFile.path);
    }
});

module.exports = router;
