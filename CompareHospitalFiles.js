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

function formatPatientNamePMD(lastName, firstName) {
    const extractFirstWord = (str) => {
        if (!str) return "";  // Handle undefined or null strings
        const words = str.split(/[ .-]/).filter(Boolean); // Split by space, dot, or dash
        return words[0] || ""; // Return the first word
    };

    return {
        lastName: extractFirstWord(lastName).toLowerCase(),
        firstName: extractFirstWord(firstName).toLowerCase()
    };
}

// Helper function to format the patient name in ECW format (Last, First)
function formatPatientNameECW(patientName) {
    if (!patientName) return { lastName: "", firstName: "" };  // Return empty if undefined or null

    const parts = patientName.split(",").map(part => part.trim());

    const extractFirstWord = (str) => {
        if (!str) return "";  // Handle undefined or null strings
        const words = str.split(/[ .-]/).filter(Boolean); // Split by space, dot, or dash
        return words[0] || ""; // Return the first word
    };

    return {  
        lastName: extractFirstWord(parts[0]).toLowerCase(),
        firstName: extractFirstWord(parts[1]).toLowerCase()
    };
}

// Helper function to compare records based on formatted names, DOB, and visit date
function compareRecords(pmdRecord, ecwRecord) {
    const pmdPatient = formatPatientNamePMD(pmdRecord["Patient Last"], pmdRecord["Patient First"]);
    const ecwPatient = formatPatientNameECW(ecwRecord["Patient"]);

    return pmdPatient.lastName === ecwPatient.lastName &&
           pmdPatient.firstName === ecwPatient.firstName &&
           pmdRecord.dob === ecwRecord.dob &&
           pmdRecord.visitDate === ecwRecord.visitDate
}


function compareExcelFiles(pmdFileBuffer, ecwFileBuffer) {
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
        charges: [record["Charge1"], record["Charge2"], record["Charge3"]].filter(Boolean),
        pmdProvider: getPmdProvider(record)
    }));

    const ecwRecords = ecwData.map(record => ({
        name: record["Patient"].toLowerCase().trim(),
        dob: isNaN(record["Patient DOB"]) ? record["Patient DOB"].trim() : formatDate(record["Patient DOB"]),
        visitDate: isNaN(record["Start Date of Service"]) ? record["Start Date of Service"].trim() : formatDate(record["Start Date of Service"]),
        cpt: record["CPT Code"],
        claimNo: record["Claim No"],
        ecwProvider: formatProvider(record["Resource Provider"])
    }));

    const matchedRecords = [];
    const missingRecords = [];
    const mistakeRecords = [];
    const duplicates = [];
    const recordTracker = new Map();

    // Map ECW records for quick lookup by name, DOB, and visitDate
    const ecwMap = {};
    ecwRecords.forEach(record => {
        const key = `${record.name}-${record.dob}-${record.visitDate}`;
        if (!ecwMap[key]) {
            ecwMap[key] = [];
        }
        ecwMap[key].push(record);
    });

    pmdRecords.forEach(pmdRecord => {
        const key = `${pmdRecord.name}-${pmdRecord.dob}-${pmdRecord.visitDate}`;
        const ecwMatches = ecwMap[key] || [];
    
        if (ecwMatches.length > 0) {
            // Process each charge from PMD
            pmdRecord.charges.forEach(charge => {
                const matchingRecords = ecwMatches.filter(ecwRecord => ecwRecord.cpt === charge);
    
                if (matchingRecords.length > 0) {
                    matchingRecords.forEach(match => {
                        const duplicateKey = `${match.name}-${match.dob}-${match.visitDate}-${charge}`;
                        
                        // Check if this combination has already been recorded
                        if (recordTracker.has(duplicateKey)) {
                            // Increment duplicate count
                            duplicates.push({
                                ...pmdRecord,
                                ecwProvider: match.ecwProvider,
                                cpt: charge,
                                claimNo: match.claimNo,
                                status: "duplicate"
                            });
                        } else {
                            // Mark the record as processed
                            recordTracker.set(duplicateKey, true);
    
                            // Check for provider mismatch
                            const providerMismatch = pmdRecord.pmdProvider !== match.ecwProvider;
                            if (providerMismatch) {
                                mistakeRecords.push({
                                    ...pmdRecord,
                                    ecwProvider: match.ecwProvider,
                                    cpt: charge,
                                    claimNo: match.claimNo,
                                    status: "provider mismatch"
                                });
                            } else {
                                matchedRecords.push({
                                    ...pmdRecord,
                                    ecwProvider: match.ecwProvider,
                                    cpt: charge,
                                    claimNo: match.claimNo,
                                    status: "matched"
                                });
                            }
                        }
                    });
                } else {
                    // Charge is missing in ECW
                    mistakeRecords.push({
                        ...pmdRecord,
                        status: `missing CPT: ${charge}`,
                        missingCPT: charge
                    });
                }
            });
        } else {
            // No matching record in ECW
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
            mistakesCount: mistakeRecords.length,
            duplicatesCount: duplicates.length
        },
        matchedRecords,
        missingRecords,
        mistakeRecords,
        duplicates
    };
}



// Helper function to format provider names
function formatProvider(provider) {
    if (!provider) return null;  // Return null if provider is undefined or null

    const parts = provider.split(",").map(part => part.trim());
    return parts.length === 2 ? `${parts[1]} ${parts[0]}`.toLowerCase() : provider.toLowerCase();
}

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
            mistakeRecords: results.mistakeRecords,
            duplicateRecords: results.duplicates 
        });
    } catch (error) {
        res.status(500).send("Error comparing Excel files: " + error.message);
    } finally {
        fs.unlinkSync(pmdFile.path);
        fs.unlinkSync(ecwFile.path);
    }
});

module.exports = router;
