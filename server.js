const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const xlsx = require('xlsx');

const app = express();
const port = 3000;

app.use(cors());
app.use(bodyParser.json());

const excelFileName = 'results.xlsx';

app.post('/save-result', (req, res) => {
    const { rollno1, name1, college1, rollno2, name2, college2, score, elapsedTime, answers } = req.body;
    console.log('Received new result:', req.body);

    const newResult = {
        'Participant 1 Roll No': rollno1,
        'Participant 1 Name': name1,
        'Participant 1 College': college1,
        'Participant 2 Roll No': rollno2,
        'Participant 2 Name': name2,
        'Participant 2 College': college2,
        'Total Score': score,
        'Time Taken (s)': elapsedTime
    };

    answers.forEach((answer, index) => {
        newResult[`Answer Q${index + 1}`] = answer;
    });

    try {
        let workbook;
        let worksheet;
        let existingData = [];
        
        try {
            workbook = xlsx.readFile(excelFileName);
            worksheet = workbook.Sheets['Results'];
            existingData = xlsx.utils.sheet_to_json(worksheet);
        } catch (error) {
            console.log('File not found, creating a new one with headers.');
            workbook = xlsx.utils.book_new();
            worksheet = xlsx.utils.json_to_sheet([newResult]);
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Results');
            xlsx.writeFile(workbook, excelFileName);
            console.log('Result saved to a new Excel file.');
            return res.json({ message: 'Result saved successfully' });
        }

        const mergedData = [...existingData, newResult];
        const newWorksheet = xlsx.utils.json_to_sheet(mergedData);
        workbook.Sheets['Results'] = newWorksheet;
        
        xlsx.writeFile(workbook, excelFileName);
        console.log('Result appended to existing Excel file.');
        res.json({ message: 'Result saved successfully' });
    } catch (error) {
        console.error('Error saving to Excel:', error);
        res.status(500).json({ message: 'Error saving result' });
    }
});

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});