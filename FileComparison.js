const express = require("express");
const router = express.Router();
const path = require("path");
const xlsx = require("xlsx");
const stringSimilarity = require("string-similarity");
const multer = require("multer");
const fs = require("fs");
var moment = require('moment');


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
    
    const formatPatientName = (name) => {
        if (!name) return { lastName: "", firstName: "" };
        
        const parts = name.split(",").map(part => part.trim());
        
        const extractFirstWord = (str) => {
            const words = str.split(/[ .-]/).filter(Boolean); // Split on space, dot, or dash
            return words[0] || ""; // Return the first word or empty string
        };

        return {  
            lastName: extractFirstWord(parts[0]).toLowerCase(),
            firstName: extractFirstWord(parts[1]).toLowerCase()
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
        const excelStartDate = new Date(1899, 11, 30);  
        const date = new Date(excelStartDate.getTime() + (serial * 86400000)); 
        return moment(date).format("MM/DD/YYYY");
    };

    const patientBillingCPTs = ["IMG1117", "IMG256778", "IMG524", "99999", "90921"];

    function cleanCPTCode(cptCode) {
        return cptCode
            .replace(/\s+/g, "") // Remove spaces
            .replace(/\(.*?\)/g, "") // Remove anything inside parentheses
            .replace(/[^0-9]/g, ""); // Keep only numeric characters
    }
    
    epicData.forEach((epicRow) => {
        if (patientBillingCPTs.includes(cleanCPTCode(epicRow["CPT Code"]))) {
            results.patientBilling.push({
                ID: epicRow.ID,
                "PatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                "SvcDate": formatDate(epicRow["Svc Date"]),
                "DOB": formatDate(epicRow["DOB"]),
                "CPTCode": cleanCPTCode(epicRow["CPT Code"]), // Store cleaned CPT code
                comment: "Record for patient billing"
            });
            results.stats.patientBillingCount = (results.stats.patientBillingCount || 0) + 1;
            return;
        }
    
        const epicName = formatPatientName(epicRow["Patient Name"]);
        const matchingRows = ecwData.filter((ecwRow) => {
            const ecwName = formatPatientName(ecwRow["Patient"]);
    
            // Exact match for first and last name
            return (
                epicName.lastName === ecwName.lastName &&
                epicName.firstName === ecwName.firstName &&
                epicRow["DOB"] === ecwRow["Patient DOB"] &&
                epicRow["Svc Date"] === ecwRow["Start Date of Service"]
            );
        });
    
        let categorized = false;
    
        if (matchingRows.length === 0) {
            results.missing.push({
                ID: epicRow.ID,
                "PatientName": formatPatientName(epicRow["Patient Name"]), 
                "SvcDate": formatDate(epicRow["Svc Date"]),
                "DOB": formatDate(epicRow["DOB"]),
                comment: "Record missing in ECW"
            });
            results.stats.missingCount++;
            categorized = true;
        } else {
            const sameCPTRows = matchingRows.filter(ecwRow => epicRow["CPT Code"] === ecwRow["CPT Code"]);
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
            }

            const duplicateRows = ecwData.filter((ecwRow) => {
                const ecwName = formatPatientName(ecwRow["Patient"]);
                
                return (
                    epicName.lastName === ecwName.lastName &&
                    epicName.firstName === ecwName.firstName &&
                    epicRow["DOB"] === ecwRow["Patient DOB"] &&
                    epicRow["Svc Date"] === ecwRow["Start Date of Service"] &&
                    epicRow[cleanCPTCode("CPT Code")] === ecwRow["CPT Code"]
                );
            });
            
            if (duplicateRows.length > 1) {
                results.duplicates.push({
                    ID: epicRow.ID,
                    "PatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                    "SvcDate": formatDate(epicRow["Svc Date"]),
                    "DOB": formatDate(epicRow["DOB"]),
                    "ecwClaimNo": duplicateRows.map(row => row["Claim No"]).join(", "),
                    comment: "Duplicate records found in ECW"
                });
                results.stats.duplicateCount++;
                categorized = true;
            } 
            else if (sameCPTRows.length > 1) {
                results.duplicates.push({
                    ID: epicRow.ID,
                    "PatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                    "SvcDate": formatDate(epicRow["Svc Date"]),
                    "DOB": formatDate(epicRow["DOB"]),
                    "ecwClaimNo": sameCPTRows.map(row => row["Claim No"]).join(", "),
                    comment: "Duplicate records found in ECW"
                });
                results.stats.duplicateCount++;
                categorized = true;
            } else if (!categorized) {
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
                    results.completelyMatched.push({
                        ID: epicRow.ID,
                        "PatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                        "EPICPatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                        "ECWPatientName": formatPatientName(ecwRow["Patient"]), // Format the name here
                        "SvcDate": formatDate(epicRow["Svc Date"]),
                        "DOB": formatDate(epicRow["DOB"]),
                        "ecwClaimNo": ecwRow["Claim No"],
                        comment: providerComparison === "exact" ? "Completely matched" : "Partially matched providers"
                    });
                    results.stats.completelyMatchedCount++;
                    categorized = true;
                } else if (!categorized) {
                    const mistakeComments = [];
                    if (!cptMatch) mistakeComments.push("CPT is incorrect");
                    missingCodes.forEach(code => mistakeComments.push(`Missing code: ${code}`));
                    if (providerComparison === "none") mistakeComments.push("Resource Provider does not match");
    
                    if (mistakeComments.length > 0) {
                        results.mistakes.push({
                            ID: epicRow.ID,
                            "PatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                           "EPICPatientName": formatPatientName(epicRow["Patient Name"]), // Format the name here
                            "ECWPatientName": formatPatientName(ecwRow["Patient"]),
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
}    

router.post("/compare", upload.fields([{ name: "epicFile" }, { name: "ecwFile" }]), async (req, res) => {
    try {
        const epicFileBuffer = fs.readFileSync(req.files["epicFile"][0].path);
        const ecwFileBuffer = fs.readFileSync(req.files["ecwFile"][0].path);
        const comparisonResults = compareExcelFiles(epicFileBuffer, ecwFileBuffer);

        res.status(200).json(comparisonResults);
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: "Error processing files", error });
    }
});

module.exports = router;
