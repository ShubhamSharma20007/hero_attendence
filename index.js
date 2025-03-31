require('dotenv').config();
const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./service-account.json');
const cors = require('cors');
const bodyParser = require('body-parser');
const { JWT } = require('google-auth-library');
const path = require('path')
const app = express();
const PORT = 5000;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.set('views', path.join(__dirname))
app.set('view engine', 'hbs')



const SPREADSHEET_ID = '1szPZ25UpLDNit1iMiaM-Yj7FrDuBu0fwik_Zx_Dz7FY';
const serviceAccountAuth = new JWT({
    email: creds.client_email,
    key: creds.private_key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

// Route to insert data
app.post('/api/add-employee', async (req, res) => {
    try {
        const { Emp_id } = req.body;
        if (!Emp_id) {
            return res.status(400).json({ message: 'Emp_id is required', success: false });
        }

        // Load the spreadsheet
        const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo();
        
        // Get the Employee_data sheet to fetch employee information
        const employeeDataSheet = doc.sheetsByTitle['Employee_data']; 
        if (!employeeDataSheet) {
            return res.status(500).json({ message: 'Employee_data sheet not found', success: false });
        }
        
        // Load rows from the Employee_data sheet
        const employeeRows = await employeeDataSheet.getRows();
        const foundEmployee = employeeRows.find(row => row._rawData[0] === Emp_id);
        if (!foundEmployee) {
            return res.status(404).json({ message: 'Invalid Employee Id', success: false });
        }
        
        const Username = foundEmployee._rawData[1];
    
        const targetSheet = doc.sheetsByIndex[0];
        
        await targetSheet.addRow({ Emp_id, Username });
        
        res.status(200).json({ message: 'Data inserted successfully!', success: true });
    } catch (error) {
        console.error('Error inserting data:', error);
        res.status(500).json({ message: 'Internal Server Error', error: error.message, success: false });
    }
});


app.use('/',(req,res)=>{
    res.render(path.join(__dirname,'/views/form.hbs'))
})

// Start Server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
