const express = require("express");
const router = express.Router();
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const moment = require("moment");
const stringSimilarity = require("string-similarity");

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

// Helper function to select the provider based on midlevel availability
function getPmdProvider(record) {
    return record["Midlevel Visit: Visit done in coordination with midlevel:"]
        ? record["Midlevel Visit: Visit done in coordination with midlevel:"].trim()
        : record["Provider"] ? record["Provider"].split(",")[0].trim().toLowerCase() : null;
}

// Helper functions for provider extraction
const getFirstNameFromResourceProvider = (resourceProvider) => {
    const trimmedResourceProvider = resourceProvider.trim().toLowerCase();
    if (trimmedResourceProvider.includes(',') && !trimmedResourceProvider.includes('-')) {
        return trimmedResourceProvider.split(",")[0].trim();
    } else if (trimmedResourceProvider.includes('-')) {
        return trimmedResourceProvider.split("-")[0].trim();
    }
};

const getLastNameFromRenderingProvider = (renderingProvider) => {
    return renderingProvider.trim().toLowerCase().split(",")[1].trim();
};

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
        pmdProvider: getPmdProvider(record)
    }));

    // Process ECW records with updated ecwProvider logic
    const ecwRecords = ecwData.map(record => {
        const resourceProviderFirstName = getFirstNameFromResourceProvider(record["Resource Provider"]);
        const renderingProviderLastName = getLastNameFromRenderingProvider(record["Rendering Provider"]);

        const ecwProviders = resourceProviderFirstName && renderingProviderLastName
            ? `${renderingProviderLastName} ${resourceProviderFirstName}`.toLowerCase()
            : null;

        return {
            name: record["Patient"].toLowerCase().trim(),
            dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
            visitDate: isNaN(record["Start Date of Service"]) ? record["Start Date of Service"].trim() : formatDate(record["Start Date of Service"]),
            cpt: record["CPT Code"],
            claimNo: record["Claim No"],
            ecwProviders
        };
    });

    // Initialize result arrays and counts
    const matchedRecords = [];
    const missingRecords = [];
    const mistakeRecords = [];
    let matchedCount = 0;
    let missingCount = 0;
    let mistakesCount = 0;

    // Create a map for ecw records with dob and visitDate as keys for quicker lookup
    const ecwMap = {};
    ecwRecords.forEach(record => {
        const key = `${record.dob}-${record.visitDate}`;
        if (!ecwMap[key]) {
            ecwMap[key] = [];
        }
        ecwMap[key].push(record);
    });

    pmdRecords.forEach(pmdRecord => {
        const key = `${pmdRecord.dob}-${pmdRecord.visitDate}`;
        const ecwMatches = ecwMap[key] || [];

        // If there are any matches, process them
        if (ecwMatches.length > 0) {
            let claimNo = null;
            const missingCPTs = [];
            let ecwProviders = null;
            let providerMismatch = false;
            let nameMatch = false;

            // Check string similarity for name comparison (above threshold 0.8)
            ecwMatches.forEach(ecwRecord => {
                const nameSimilarity = stringSimilarity.compareTwoStrings(pmdRecord.name, ecwRecord.name);
                
                // If similarity is above 0.8 and both visit date and dob match, count as a match
                if (nameSimilarity > 0.8 && pmdRecord.dob === ecwRecord.dob && pmdRecord.visitDate === ecwRecord.visitDate) {
                    nameMatch = true;
                    claimNo = ecwRecord.claimNo;
                    ecwProviders = ecwRecord.ecwProviders;
                }
            });

            if (nameMatch) {
                pmdRecord.charges.forEach(charge => {
                    const match = ecwMatches.find(ecwRecord => ecwRecord.cpt === charge);
                    if (match) {
                        claimNo = match.claimNo;
                        ecwProviders = match.ecwProviders;
                    } else {
                        missingCPTs.push(charge);
                    }
                });

                const providerToCompare = pmdRecord.pmdProvider;
                if (providerToCompare && ecwProviders && providerToCompare !== ecwProviders) {
                    providerMismatch = true;
                }

                if (missingCPTs.length > 0 || providerMismatch) {
                    mistakeRecords.push({
                        ...pmdRecord,
                        status: `${missingCPTs.length > 0 ? `missing CPTs: ${missingCPTs.join(", ")}` : ""}${providerMismatch ? " Billing providers not matched" : ""}`,
                        missingCPTs,
                        claimNo,
                        ecwProviders
                    });
                } else {
                    matchedRecords.push({
                        ...pmdRecord,
                        status: "matched",
                        charges: pmdRecord.charges,
                        claimNo,
                        ecwProviders
                    });
                }
            } else {
                missingRecords.push({
                    ...pmdRecord,
                    status: "missing",
                    charges: pmdRecord.charges
                });
            }
        } else {
            missingRecords.push({
                ...pmdRecord,
                status: "missing",
                charges: pmdRecord.charges
            });
        }
    });

    return {
        stats: {
            matchedCount: matchedRecords.length,
            missingCount: missingRecords.length,
            mistakesCount: mistakeRecords.length
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
        fs.unlinkSync(pmdFile.path);
        fs.unlinkSync(ecwFile.path);
    }
});

module.exports = router;
