const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const path = require('path');
const app = express();

// Configure Express
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine', 'ejs');

// Your Google Sheets API Key and Sheet ID
const apiKey = 'AIzaSyBetGkBUBGg8vqxEYAnwm0FMXUCrole-xs'; // Replace with your Google Sheets API Key
const spreadsheetId = '1QqxjcZj8YZWHNri6quC3Cu3uUo9pWrlP9P73k3oKUAM'; // Replace with your Google Sheets ID

// Route to Render UI
app.get('/', (req, res) => {
  res.render('index');
});

// Route to Handle Data Export
app.post('/export', async (req, res) => {
  try {
    const range = 'Sheet1!A:I'; // Adjust if needed
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?key=${apiKey}`;
    const response = await axios.get(url);

    const data = response.data.values;
    const [headers, ...rows] = data;

    const filteredData = rows.filter(row => {
      const [agent, department, region, , date] = row;
      const isDateMatch = date === req.body.date;
      const isDepartmentMatch = department === req.body.department;
      const isRegionMatch = region === req.body.region;
      const isAgentMatch = agent === req.body.agent;
      return isDateMatch && isDepartmentMatch && isRegionMatch && isAgentMatch;
    });

    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([headers, ...filteredData]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Data');

    // Summary Report
    const summary = filteredData.reduce((acc, row) => {
      const [agent, department, region, , date, , , taskTime] = row;
      if (!acc[date]) acc[date] = {};
      if (!acc[date][agent]) acc[date][agent] = { count: 0, totalTime: 0 };

      acc[date][agent].count += 1;
      acc[date][agent].totalTime += parseFloat(taskTime) || 0;

      return acc;
    }, {});

    const summaryData = [['Agent Name', 'Working Department', 'Working Region', 'Total Task Time', 'Submission Count', 'Date']];
    Object.keys(summary).forEach(date => {
      Object.keys(summary[date]).forEach(agent => {
        const { count, totalTime } = summary[date][agent];
        summaryData.push([agent, '', '', totalTime, count, date]);
      });
    });

    const summarySheet = xlsx.utils.aoa_to_sheet(summaryData);
    xlsx.utils.book_append_sheet(workbook, summarySheet, 'Summary Report');

    const filePath = path.join(__dirname, 'exported-data.xlsx');
    xlsx.writeFile(workbook, filePath);

    res.download(filePath, 'data.xlsx', () => {
      // Clean up file after download
      fs.unlinkSync(filePath);
    });
  } catch (error) {
    res.status(500).send('Error exporting data');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
