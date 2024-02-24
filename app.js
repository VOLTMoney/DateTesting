const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');
const {convertToTimestamp} = require("./utils");

const app = express();
const port = 3000;

// Middleware to parse JSON in the request body
app.use(express.json());

// Middleware to handle URL-encoded form data
app.use(express.urlencoded({ extended: true }));

// Serve static files from the 'public' directory
app.use(express.static('public'));

// Sample route
app.get('/', (req, res) => {

    const apiUrl = 'http://internal-Alpha-Credi-12R5GMOMSALNO-1994618058.ap-south-1.elb.amazonaws.com/disbursal/testing';
    const workbookRead = new ExcelJS.Workbook();
    const worksheetRead = workbookRead.addWorksheet('Sheet1');

    const workbookWrite = new ExcelJS.Workbook();
    const worksheetWrite = workbookWrite.addWorksheet('Sheet1');

    worksheetWrite.columns = [
        { header: 'Request Date', key: 'requestDate', width: 10 },
        { header: 'Response Date', key: 'responseDate', width: 50 },
    ];

    async function makeAPICallAndSave(day) {
        try {
            const requestBody = {
                "creditId": "8a8013bd8d7e1c0c018d8253fed6003d",
                "dateTime": `${day}`
            };

            const response = await axios.post(apiUrl, requestBody, {
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            // Save the response to the Excel sheet
            console.log('requestDate: ', day);
            console.log('response: ', JSON.stringify(response.data.expectedDisbursalTransafer));

            worksheetWrite.addRow({
                requestDate: day,
                responseDate: JSON.stringify(response.data.expectedDisbursalTransafer)
            });

            console.log(`Day ${day} - Response saved.`);
            // Introduce a 3-second delay
            await new Promise(resolve => setTimeout(resolve, 250));
        } catch (error) {
            console.error(`Day ${day} - Error:`, error.message);
        }
    }

    async function makeAPICallsFor365Days(timeArray) {
        for (let i = 0; i <= timeArray.length - 1; i++) {
            await makeAPICallAndSave(timeArray[i]);
        }
        // Save the Excel file
        await workbookWrite.xlsx.writeFile('api_responses1.xlsx');
        console.log('Excel file saved.');
    }

    const convertToTimestamp = (time) => {
        const originalTimestamp = time;
        const dateObject = new Date(originalTimestamp);
        const formattedDate = dateObject.toISOString().replace("T", " ").replace(".000Z", "");
        return formattedDate
    }

    workbookRead.xlsx.readFile('auto_template.xlsx')
        .then(async () => {
            // Assuming the sheet name is 'API Responses'
            const worksheet = workbookRead.getWorksheet('Sheet1');

            let timeArray = [];

            // Iterate through rows and access cell values
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                // const day = row.getCell(o).value;
                let timeStamp = row.getCell(1).value;
                timeStamp = convertToTimestamp(timeStamp);
                timeArray.push(timeStamp);
            });

            timeArray.map((day)=> {
                console.log("timeArray: ", day);
            })
            await makeAPICallsFor365Days(timeArray);
        })
        .catch(error => {
            console.error('Error reading Excel file:', error.message);
        });

        return res.send('Hello')
});

// Sample route with a route parameter
app.get('/user/:id', (req, res) => {
    const creditId = req.params.id;

    const apiUrl = 'http://internal-Alpha-Credi-12R5GMOMSALNO-1994618058.ap-south-1.elb.amazonaws.com/disbursal/testing';
    const workbookRead = new ExcelJS.Workbook();
    const worksheetRead = workbookRead.addWorksheet('Sheet1');

    const workbookWrite = new ExcelJS.Workbook();
    const worksheetWrite = workbookWrite.addWorksheet('Sheet1');

    worksheetWrite.columns = [
        { header: 'Request Date', key: 'requestDate', width: 10 },
        { header: 'Response Date', key: 'responseDate', width: 50 },
    ];

    async function makeAPICallAndSave(day) {
        try {
            const requestBody = {
                // "creditId": "8a8013bd8d7e1c0c018d8253fed6003d",
                "creditId": `${creditId}`,
                "dateTime": `${day}`
            };

            const response = await axios.post(apiUrl, requestBody, {
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            // Save the response to the Excel sheet
            console.log('requestDate: ', day);
            console.log('response: ', JSON.stringify(response.data.expectedDisbursalTransafer));

            worksheetWrite.addRow({
                requestDate: day,
                responseDate: response.data.expectedDisbursalTransafer
            });

            console.log(`Day ${day} - Response saved.`);
            // Introduce a 3-second delay
            await new Promise(resolve => setTimeout(resolve, 250));
        } catch (error) {
            console.error(`Day ${day} - Error:`, error.message);
        }
    }

    async function makeAPICallsFor365Days(timeArray) {
        for (let i = 0; i <= timeArray.length - 1; i++) {
            await makeAPICallAndSave(timeArray[i]);
        }
        // Save the Excel file
        await workbookWrite.xlsx.writeFile('api_responses1.xlsx');
        console.log('Excel file saved.');
    }

    const convertToTimestamp = (time) => {
        const originalTimestamp = time;
        const dateObject = new Date(originalTimestamp);
        const formattedDate = dateObject.toISOString().replace("T", " ").replace(".000Z", "");
        return formattedDate
    }

    workbookRead.xlsx.readFile('auto_template.xlsx')
        .then(async () => {
            // Assuming the sheet name is 'API Responses'
            const worksheet = workbookRead.getWorksheet('Sheet1');

            let timeArray = [];

            // Iterate through rows and access cell values
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                // const day = row.getCell(o).value;
                let timeStamp = row.getCell(1).value;
                timeStamp = convertToTimestamp(timeStamp);
                timeArray.push(timeStamp);
            });

            timeArray.map((day)=> {
                console.log("timeArray: ", day);
            })
            await makeAPICallsFor365Days(timeArray);
        })
        .catch(error => {
            console.error('Error reading Excel file:', error.message);
        });

    res.send(`User ID: ${userId} report generated`);
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).send('Something went wrong!');
});

// Start the server
app.listen(port, () => {
    console.log(`Server is listening on port ${port}`);
});
