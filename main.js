// Import excel JS package
const Excel = require('exceljs');

async function parseData() {
    // Read from a file
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("data.xlsx");
    const outputWorksheet = new Excel.Workbook();

    // Prep working vars
    let date = null;
    let humDataByDay = [];
    let tempDataByDay = [];

    // Loop through each tab/sheet
    workbook.eachSheet((ws, wId) => {
        console.log("Parsing: ", ws.name);
        let outputByDay = [];
        date = null;

        // Loop through each row of each tab/sheet
        ws.eachRow((row, i) => {
            // Skip the first row since that's the header text
            if(i > 1) {
                // Dates in javascript are notoriously a pain to work with, so standardize the date before processing further
                if(date == null) {
                    date = row.getCell(1).value;
                    date = new Date(`${date.toISOString().substring(0,10)} MST`);
                    date = `${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()}`
                }
                let rowDate = row.getCell(1).value;
                rowDate = new Date(`${rowDate.toISOString().substring(0,10)} MST`);
                rowDate = `${rowDate.getMonth()+1}/${rowDate.getDate()}/${rowDate.getFullYear()}`

                // Same day, different time?
                if(date == rowDate) {
                    humDataByDay.push(row.getCell(3).value);
                    tempDataByDay.push(row.getCell(4).value);
                    
                    // Last cell? This should prob be added to function so I'm not repeating myself later
                    if (i == (ws.rowCount - 1)) {
                        let humMin = Math.min(...humDataByDay);
                        let humMax = Math.max(...humDataByDay);
                        let humMean = ((humDataByDay.reduce((a,c) => a + c, 0))/(humDataByDay.length))
                        let tempMin = Math.min(...tempDataByDay);
                        let tempMax = Math.max(...tempDataByDay);
                        let tempMean = ((tempDataByDay.reduce((a,c) => a + c, 0))/(tempDataByDay.length))
                        outputByDay.push({
                            date: date,
                            meanT: tempMean,
                            maxT: tempMax,
                            minT: tempMin,
                            meanHum: humMean,
                            maxHum: humMax,
                            minHum: humMin
                        });
    
                        // Clear out data from previous day
                        humDataByDay = [];
                        tempDataByDay = [];
    
                        // Then start adding new data
                        humDataByDay.push(row.getCell(3).value);
                        tempDataByDay.push(row.getCell(4).value);
    
                        
                        date = rowDate;
                    }
                // New day 
                } else {
                    // Perform calculations
                    let humMin = Math.min(...humDataByDay);
                    let humMax = Math.max(...humDataByDay);
                    let humMean = ((humDataByDay.reduce((a,c) => a + c, 0))/(humDataByDay.length))
                    let tempMin = Math.min(...tempDataByDay);
                    let tempMax = Math.max(...tempDataByDay);
                    let tempMean = ((tempDataByDay.reduce((a,c) => a + c, 0))/(tempDataByDay.length))
                    outputByDay.push({
                        date: date,
                        meanT: tempMean,
                        maxT: tempMax,
                        minT: tempMin,
                        meanHum: humMean,
                        maxHum: humMax,
                        minHum: humMin
                    });

                    // Clear out data from previous day
                    humDataByDay = [];
                    tempDataByDay = [];

                    // Then start adding new data
                    humDataByDay.push(row.getCell(3).value);
                    tempDataByDay.push(row.getCell(4).value);

                    
                    date = rowDate;
                }
            } 
        });

        // Write data to new .xlsx file
        let sheet = outputWorksheet.addWorksheet(ws.name);
        sheet.columns = [
            { header: 'Date', key: 'date' },
            { header: 'MeanT', key: 'meanT' },
            { header: 'MaxT', key: 'maxT' },
            { header: 'MinT', key: 'minT' },
            { header: 'MeanHum', key: 'meanHum' },
            { header: 'MaxHum', key: 'maxHum' },
            { header: 'MinHum', key: 'minHum' },
        ];
        sheet.addRows(outputByDay);
    });
    
    await outputWorksheet.xlsx.writeFile("test.xlsx");

}

parseData();