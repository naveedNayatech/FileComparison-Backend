const express = require('express');
const app = express();
const compareApi = require('./FileComparison'); // Import the API
const cors = require('cors');

const port = 3000;

app.use(cors({
  origin: 'http://localhost:5173', // Replace with your frontend's origin
}));

// Use the API routes
app.use('/api', compareApi); // Use the compare API under '/api' path

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
