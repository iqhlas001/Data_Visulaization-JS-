// Import the necessary libraries
const express = require('express'); // Express framework for handling HTTP requests
const xlsx = require('xlsx'); // Library for reading Excel files
const cors = require('cors');
const port = 4000; // Define the port on which the server will run
const app = express();
app.use(cors());
// app.use(cors({
//     origin: '*', // Change this to specific origins for production
// }));

// Function to convert Excel serial date to JavaScript Date
const excelDateToJSDate = (serial) => {
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
};

// Helper function to format date to "DD-MM-YYYY"
const formatDate = (date) => {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Month is 0-based
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
};

// Read the Excel file and convert to JSON
const getDataFromExcel = () => {
    const workbook = xlsx.readFile('dummy.xlsx'); // Load your Excel file
    const sheetName = workbook.SheetNames[0]; // Get the name of the first sheet
    const sheet = workbook.Sheets[sheetName]; // Get the sheet object
    return xlsx.utils.sheet_to_json(sheet); // Parse the sheet to JSON
}

//----------------TESTING-------------------------

const loadAndProcessData = () => {
    const workbook = xlsx.readFile('dummy.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    return data.map(item => ({
        ...item,
        SSL: item.Ordered > 0 ? (item.Recv / item.Ordered).toFixed(2) : 0,
        Date: formatDate(excelDateToJSDate(item.Date)) // Convert Excel date serial to JS Date
    }));
};



app.get('/api/pivot', (req, res) => {
    try {
        const data = loadAndProcessData();
        const {year, week, date, country, div, supplierNo, supplierName, poNo } = req.query;

        // Filter data based on query parameters
        let filteredData = data.filter(item => {
            const itemDate = new Date(item.Date); // Assuming item.Date converted to a Date object
            return (!year || item.Year === year) &&
                   (!week || item.Week === Number(week)) &&
                   (!date || item.Date === Number(date)) &&
                   (!country || item.country === country) &&
                   (!div || item.div === div) &&
                   (!supplierNo || item['Supplier NO'] === Number(supplierNo)) &&
                   (!supplierName || item['Supplier Name'] === supplierName) &&
                   (!poNo || item['PO NO'] === Number(poNo))

        });

        res.json(filteredData);
    } catch (error) {
        console.error('Error loading data:', error);
        res.status(500).json({ error: 'Internal Server Error' });
    }
});

   
//----------------TESTING-------------------------

// API to get raw data (all)
app.get('/api/data', (req, res) => {
    const data = getDataFromExcel()
    res.json(data); // Response with the data
});

// Get countries (Only)
app.get('/api/data/countries', (req, res) => {
    const data = getDataFromExcel();
    const countries = [...new Set(data.map(item => item.country))];
    res.json(countries);
});

// Get divisions (Only)
app.get('/api/data/divisions', (req, res) => {
    const data = getDataFromExcel();
    const divisions = [...new Set(data.map(item => item.div))]
    res.json(divisions);
});


// API to calculate SSL with all data showing
// app.get('/api/ssl', (req, res) => {
//     const data = getDataFromExcel();
//     const calculatedData = data.map(item => ({
//         ...item,
//         SSL: item.Ordered > 0 ? (item.Recv / item.Ordered) : 0,
//     }));
//     res.json(calculatedData);
// });

// API to calculate SSL each row
app.get('/api/ssl1', (req, res) => {
    const data = getDataFromExcel();
    const sslData = data.map(item => ({
        SSL: item.Ordered > 0 ? (item.Recv / item.Ordered) : 0,
    }));
    res.json(sslData);
});

// API to calculate total SSL by division
app.get('/api/total-ssl', (req, res) => {
    const data = getDataFromExcel();
    const totalSSL = data.reduce((acc, item) => {
        const ssl = item.Ordered > 0 ? (item.Recv / item.Ordered) : 0;
        acc[item.div] = (acc[item.div] || 0) + ssl;
        return acc;
    }, {});
    res.json(totalSSL);
});

//API to calculate total SSL by country
app.get('/api/countryssl', (req, res) => {
    const { country, div } = req.query; // Extract country and div from query parameters
    const data = getDataFromExcel();
    
    // Filter data based on selected country and div
    const filteredData = data.filter(item => 
        (!country || item.country === country) && 
        (!div || item.div === div)
    );

    const totalSSL = filteredData.reduce((acc, item) => {
        const ssl = item.Ordered > 0 ? (item.Recv / item.Ordered) : 0;
        acc[item.country] = (acc[item.country] || 0) + ssl; // Accumulate by country
        return acc;
    }, {});

    res.json(totalSSL);
});

app.listen(port, () => { // Start the server and listen on the defined port
    console.log(`Server running on http://localhost:${port}`); // Log the URL to the console
});

// app.use(express.static('public')); // Create a 'public' directory for static files