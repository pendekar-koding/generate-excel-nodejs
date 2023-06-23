const express = require('express');
const ExcelJS = require('exceljs');
const app = express();
const port = 3000;

// Needed to parse JSON data from the request body
app.use(express.json());

app.post('/generate-excel', async (req, res) => {
    // Create a new instance of a Workbook class
    let workbook = new ExcelJS.Workbook();

    // Add Worksheets to the workbook
    let worksheet = workbook.addWorksheet('Sheet 1');

    // Add Headers to the Worksheet
    const headers = req.body.headers.map((header) => {
        return { header: header, key: header, width:30 };
    });
    worksheet.columns = headers;


    // Add Rows to the Worksheet
    const data = req.body.data;
    data.forEach((row) => {
        console.log(row);
        worksheet.addRow(row);
    });

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "Report.xlsx"
    );

    // Write to response
    await workbook.xlsx.write(res);

    res.end();
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
