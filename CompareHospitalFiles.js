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

function getPmdProvider(record) {
    return record["Midlevel Visit: Visit done in coordination with midlevel:"]
        ? record["Midlevel Visit: Visit done in coordination with midlevel:"].toLowerCase().trim().split(" ")[0] +" - "+
        record["Provider"]
            .split(",")[0]
            .trim()
            .split(" ")
            .pop()
            .toLowerCase()
        : record["Provider"] ? record["Provider"].split(",")[0].trim().toLowerCase() : null;
}

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
           pmdRecord.visitDate === ecwRecord.visitDate;
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
    
        if (ecwMatches.length > 0) {
            const missingCPTs = [];
            let duplicateDetected = false;

    
            pmdRecord.charges.forEach(charge => {
                const matchingRecords = ecwMatches.filter(ecwRecord =>
                    ecwRecord.cpt === charge &&
                    compareRecords(pmdRecord, ecwRecord) // Compare based on names, DOB, and visit date
                );
                
                if (matchingRecords.length > 0) {
                    matchingRecords.forEach(match => {
                        const duplicateKey = `${match.name}-${match.dob}-${match.visitDate}-${match.cpt}`;
                        if (recordTracker.has(duplicateKey)) {
                            duplicateDetected = true;
                            duplicates.push({
                                ...pmdRecord,
                                ecwProvider: match.ecwProvider,
                                cpt: charge,
                                claimNo: match.claimNo,
                                status: "duplicate"
                            });
                        } else {
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
                                recordTracker.set(duplicateKey, true);
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
                    // If no match found, add to missingCPTs
                    missingCPTs.push(charge);
                }
            });
    
            // If there are any missing CPTs, flag this record as a mistake
            if (missingCPTs.length > 0) {
                mistakeRecords.push({
                    ...pmdRecord,
                    status: `missing CPTs: ${missingCPTs.join(", ")}`,
                    missingCPTs
                });
            }
        } else {
            // If no matching records from ECW, add this to missingRecords
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


function formatProvider(provider) {
    if (!provider) return null;

    const parts = provider.split(",").map(part => part.trim());
    return parts.length === 2 ? `${parts[1]} ${parts[0]}`.toLowerCase().trim() : provider.toLowerCase().trim();
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
