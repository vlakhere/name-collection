const express = require('express');
const bodyParser = require('body-parser');
const excel = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.use(express.static('public'));

app.post('/submit', async (req, res) => {
    const { name } = req.body;

    if (!name) {
        return res.status(400).send('Name is required.');
    }

    try {
        // Load the existing workbook or create a new one
        const workbookPath = path.join(__dirname, 'users.xlsx');
        const workbook = new excel.Workbook();
        let worksheet;

        try {
            await workbook.xlsx.readFile(workbookPath);
            worksheet = workbook.getWorksheet(1);
        } catch (error) {
            worksheet = workbook.addWorksheet('Users');
            worksheet.addRow(['Name']);
        }

        // Add the user's name to the worksheet
        worksheet.addRow([name]);

        // Save the workbook
        await workbook.xlsx.writeFile(workbookPath);

        res.status(200).send('Name successfully added to the Excel file.');
    } catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
});

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
