const express = require('express');
const app = express();
const FileComparison = require('./FileComparison');
const compareHospitalFiles = require('./CompareHospitalFiles');
const PMDMissingCPTChecking = require('./PMDMissingCPTChecking');
const PMDDuplicateCheckRecord = require('./PMDDuplicateCheck');
const cors = require('cors');

const port = 3000;

app.use(cors({
  origin: 'http://localhost:5173', // Replace with your frontend's origin
}));

// Use the API routes
app.use('/api', FileComparison);
app.use('/api/hospital', compareHospitalFiles)
app.use('/api/pmd',PMDMissingCPTChecking);
app.use('/api/duplicate', PMDDuplicateCheckRecord);

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
