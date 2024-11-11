const express = require("express");
const router = express.Router();
const path = require("path");
const xlsx = require("xlsx");
const stringSimilarity = require("string-similarity");
const multer = require("multer");
const fs = require("fs");
var moment = require('moment');


// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, './uploads'); // Ensure you have an "uploads" folder
    },
    filename: (req, file, cb) => {
      cb(null, Date.now() + path.extname(file.originalname)); // Append timestamp to file name
    }
});
const upload = multer({ storage: storage });
  

const compareExcelFiles = (epicFileBuffer, ecwFileBuffer) => {

    const epicFile = xlsx.read(epicFileBuffer, { type: "buffer" });
    const ecwFile = xlsx.read(ecwFileBuffer, { type: "buffer" });

    const epicSheet = epicFile.Sheets[epicFile.SheetNames[0]];
    const ecwSheet = ecwFile.Sheets[ecwFile.SheetNames[0]];

    const epicData = xlsx.utils.sheet_to_json(epicSheet);
    const ecwData = xlsx.utils.sheet_to_json(ecwSheet);

    console.log('Generating a result');

    const results = {
        completelyMatched: [],
        missing: [],
        duplicates: [],
        mistakes: [],
        patientBilling: [],
        stats: {
            completelyMatchedCount: 0,
            missingCount: 0,
            duplicateCount: 0,
            mistakeCount: 0,
            patientBillingCount: 0
        },
    };

    const extractDiagnosisCodes = (diagnosis) => 
        (diagnosis?.match(/\[(.*?)\]/g) || []).map(code => code.replace(/[\[\]]/g, "")).slice(0, 2);

    const formatName = (name) => {
        if (!name) return { lastName: "", firstName: "" };
        const parts = name.split(",").map(part => part.trim());
        return {
            lastName: parts[0] || "",
            firstName: parts[1] || ""
        };
    };

    const formatResourceProvider = (provider) => {
        if (!provider) return { firstName: "", lastName: "" };
        const parts = provider.split("-").map(part => part.trim().replace(/,$/, ""));
        return {
            firstName: parts[0] || "",
            lastName: parts[1] || ""
        };
    };


    const providerMatch = (epicProviderFirstName, epicProviderLastName, ecwProvider) => {
        const epicProviderFormatted = `${epicProviderFirstName} ${epicProviderLastName}`.toLowerCase().replace(/[-,]/g, '').trim();
        const ecwProviderFormatted = `${ecwProvider.firstName} ${ecwProvider.lastName}`.toLowerCase().replace(/[-,]/g, '').trim();
    
        if (epicProviderFormatted === ecwProviderFormatted) return "exact";
    
        // Lower similarity threshold to 0.7
        const similarityScore = stringSimilarity.compareTwoStrings(epicProviderFormatted, ecwProviderFormatted);
        return similarityScore >= 0.7 ? "partial" : "none";
    };

    const formatDate = (serial) => {
        // Excel serial date starts from 1899-12-30, so we add the serial days (multiply by 86400000 for milliseconds in a day)
        const excelStartDate = new Date(1899, 11, 30);  // Starting date of Excel's serial date system
        const date = new Date(excelStartDate.getTime() + (serial * 86400000)); // Add serial days to starting date
    
        return moment(date).format("MM/DD/YYYY");
    };

    // Define CPT codes for patientBilling category
    const patientBillingCPTs = ["IMG1117", "IMG256778", "IMG524", "99999"];

    epicData.forEach((epicRow) => {
        // Check if the CPT code belongs to patientBilling category
        if (patientBillingCPTs.includes(epicRow["CPT Code"])) {
            results.patientBilling.push({
                ID: epicRow.ID,
                "PatientName": epicRow["Patient Name"],
                "SvcDate": formatDate(epicRow["Svc Date"]),
                "DOB": formatDate(epicRow["DOB"]),
                "CPTCode": epicRow["CPT Code"],
                comment: "Record for patient billing"
            });
            results.stats.patientBillingCount = (results.stats.patientBillingCount || 0) + 1;
            return; // Skip further processing for this row
        }
    
        // Proceed with original matching logic as before...
        const epicName = formatName(epicRow["Patient Name"]);
        const matchingRows = ecwData.filter((ecwRow) => {
            const ecwName = formatName(ecwRow["Patient"]);
    
            return (
                ecwName.lastName === epicName.lastName &&
                ecwName.firstName === epicName.firstName &&
                ecwRow["Patient DOB"] === epicRow["DOB"] &&
                ecwRow["Start Date of Service"] === epicRow["Svc Date"]
            );
        });
    
        let categorized = false; // Flag to prevent multiple categorizations
    
        if (matchingRows.length === 0) {
            // No match found, mark as missing
            results.missing.push({
                ID: epicRow.ID,
                "PatientName": epicRow["Patient Name"],
                "SvcDate": formatDate(epicRow["Svc Date"]),
                "DOB": formatDate(epicRow["DOB"]),
                comment: "Record missing in ECW"
            });
            results.stats.missingCount++;
            categorized = true;
        } else {
            // Check if matching rows have identical or different CPT codes
            const sameCPTRows = matchingRows.filter(ecwRow => ecwRow["CPT Code"] === epicRow["CPT Code"]);
            const differentCPTRows = matchingRows.length > 1 && sameCPTRows.length === 0;
    
            if (differentCPTRows) {
                // If CPT codes differ, count as matched instead of duplicate
                results.completelyMatched.push({
                    ID: epicRow.ID,
                    "PatientName": epicRow["Patient Name"],
                    "SvcDate": formatDate(epicRow["Svc Date"]),
                    "DOB": formatDate(epicRow["DOB"]),
                    "ecwClaimNo": matchingRows.map(row => row["Claim No"]).join(", "),
                    comment: "Matched in ECW with different CPT codes"
                });
                results.stats.completelyMatchedCount++;
                categorized = true;
            } else if (sameCPTRows.length > 1) {
                // Count as duplicate if multiple rows match exactly on CPT code
                results.duplicates.push({
                    ID: epicRow.ID,
                    "PatientName": epicRow["Patient Name"],
                    "SvcDate": formatDate(epicRow["Svc Date"]),
                    "DOB": formatDate(epicRow["DOB"]),
                    "ecwClaimNo": sameCPTRows.map(row => row["Claim No"]).join(", "),
                    comment: "Duplicate records found in ECW"
                });
                results.stats.duplicateCount++;
                categorized = true;
            } else if (!categorized) {
                // Check for complete match conditions
                const ecwRow = sameCPTRows[0] || matchingRows[0];
                const epicCPT = epicRow["CPT Code"].split(" ")[0];
                const ecwCPT = ecwRow["CPT Code"].toString();
                const cptMatch = epicCPT === ecwCPT;
    
                const epicDiagnosisCodes = extractDiagnosisCodes(epicRow["Diagnosis"]);
                const icdCodes = [
                    ecwRow["ICD1 Code"],
                    ecwRow["ICD2 Code"],
                    ecwRow["ICD3 Code"],
                    ecwRow["ICD4 Code"]
                ];
                const missingCodes = epicDiagnosisCodes.filter(code => !icdCodes.includes(code));
    
                const epicProviderFirstName = formatName(epicRow["Service Provider"]).firstName;
                const epicProviderLastName = formatName(epicRow["Billing Provider"]).lastName;
                const ecwProvider = formatResourceProvider(ecwRow["Resource Provider"]);
                
                const providerComparison = providerMatch(epicProviderFirstName, epicProviderLastName, ecwProvider);
    
                if ((providerComparison === "exact" || providerComparison === "partial") && cptMatch && missingCodes.length === 0) {
                    // Count it as completely matched
                    results.completelyMatched.push({
                        ID: epicRow.ID,
                        "PatientName": epicRow["Patient Name"],
                        "SvcDate": formatDate(epicRow["Svc Date"]),
                        "DOB": formatDate(epicRow["DOB"]),
                        "ecwClaimNo": ecwRow["Claim No"],
                        comment: providerComparison === "exact" ? "Completely matched" : "Partially matched providers"
                    });
                    results.stats.completelyMatchedCount++;
                    categorized = true;
                } else if (!categorized) {
                    // Otherwise, count as mistake with specified comments
                    const mistakeComments = [];
                    if (!cptMatch) mistakeComments.push("CPT is incorrect");
                    missingCodes.forEach(code => mistakeComments.push(`Missing code: ${code}`));
                    if (providerComparison === "none") mistakeComments.push("Resource Provider does not match");
    
                    if (mistakeComments.length > 0) {
                        results.mistakes.push({
                            ID: epicRow.ID,
                            "PatientName": epicRow["Patient Name"],
                            "SvcDate": formatDate(epicRow["Svc Date"]),
                            "DOB": formatDate(epicRow["DOB"]),
                            "EPICProvider": `${epicProviderFirstName} ${epicProviderLastName}`,
                            "ECWProvider": `${ecwProvider.firstName} ${ecwProvider.lastName}`,
                            "ecwClaimNo": ecwRow["Claim No"],
                            comment: mistakeComments.join("; "),
                        });
                        results.stats.mistakeCount++;
                    }
                }
            }
        }
    });    
    
    return results;
};

router.post("/compare", upload.fields([{ name: "epicFile" }, { name: "ecwFile" }]), (req, res) => {
   
    const epicFile = req.files['epicFile'] ? req.files['epicFile'][0] : null;
    const ecwFile = req.files['ecwFile'] ? req.files['ecwFile'][0] : null;

    if (!epicFile || !ecwFile) {
        return res.status(400).json({ error: 'Both files are required' });
    }

    const epicFileBuffer = fs.readFileSync(epicFile.path);
    const ecwFileBuffer = fs.readFileSync(ecwFile.path);

    try {
    
        const results = compareExcelFiles(epicFileBuffer, ecwFileBuffer);

        res.json({
            message: "Comparison complete",
            results,
        });
    } catch (error) {
        res.status(500).send("Error comparing Excel files: " + error.message);
    }
});


// router.get("/compare/missing", (req, res) => {
//     try {
//         const results = compareExcelFiles();
//         res.json({
//             totalMissing: results.stats.missingCount,
//             missingRecords: results.missing,
//         });
//     } catch (error) {
//         res.status(500).send("Error retrieving missing records: " + error.message);
//     }
// });

// router.get("/compare/completelyMatched", (req, res) => {
//     try {
//         const results = compareExcelFiles();
//         res.json({
//             totalCompletelyMatched: results.stats.completelyMatchedCount,
//             completelyMatchedRecords: results.completelyMatched,
//         });
//     } catch (error) {
//         res.status(500).send("Error retrieving completely matched records: " + error.message);
//     }
// });

// router.get("/compare/partiallyMatched", (req, res) => {
//     try {
//         const results = compareExcelFiles();
//         res.json({
//             totalPartiallyMatched: results.stats.partiallyMatchedCount,
//             partiallyMatchedRecords: results.partiallyMatched,
//         });
//     } catch (error) {
//         res.status(500).send("Error retrieving partially matched records: " + error.message);
//     }
// });

// router.get("/compare/duplicates", (req, res) => {
//     try {
//         const results = compareExcelFiles();
//         res.json({
//             totalDuplicates: results.stats.duplicateCount,
//             duplicateRecords: results.duplicates,
//         });
//     } catch (error) {
//         res.status(500).send("Error retrieving duplicate records: " + error.message);
//     }
// });

// router.get("/compare/mistakes", (req, res) => {
//     try {
//         const results = compareExcelFiles();
//         res.json({
//             totalMistakes: results.stats.mistakeCount,
//             mistakeRecords: results.mistakes,
//         });
//     } catch (error) {
//         res.status(500).send("Error retrieving mistakes: " + error.message);
//     }
// });

module.exports = router;
