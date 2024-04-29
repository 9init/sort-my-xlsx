const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');
const { exec } = require('child_process');
const natsort = require('natsort').default();

// enums for OS, linux, mac, windows
const OS = {
    LINUX: 'linux',
    MAC: 'darwin',
    WINDOWS: 'win32'
};

// Function to sort filenames and class values (case-insensitive)
function sortFileData(data, os = OS.WINDOWS) {
    const firstColumnName = Object.keys(data[0])[0];
    return data.sort((a, b) => {
        const imgA = a[firstColumnName];
        const imgB = b[firstColumnName];
        if (!imgA || !imgB) {
            // Handle cases where the filename is missing
            return 0;
        }

        let filenameA = imgA.toLowerCase()
        let filenameB = imgB.toLowerCase()

        filenameA = filenameA.replace(/[^a-zA-Z0-9]/g, c => c.charCodeAt(0));
        filenameB = filenameB.replace(/[^a-zA-Z0-9]/g, c => c.charCodeAt(0));

        // if (os === OS.WINDOWS) {
        //     // replace all characters to its equivalent ascii value

        //     // replace all characters to its equivalent ascii value
        //     // filenameA = filenameA.replace(/[^a-zA-Z0-9]/g, ' ');
        //     // filenameB = filenameB.replace(/[^a-zA-Z0-9]/g, ' ');
        // }
        
        return natsort(filenameA, filenameB);
    });
}

function ls(dir) {
    return new Promise((resolve, reject) => {
        exec(`exa -1 ${dir}`, (error, stdout, stderr) => {
            if (error) {
                reject(error);
                return;
            }
            if (stderr) {
                reject(stderr);
                return;
            }
            resolve(stdout.split('\n').filter(Boolean));
        });
    })
}

// Function to read data from Excel file (assuming it's a single sheet)
function readExcelData(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet)
    return data
}

// Function to write data to a new Excel file
function writeExcelData(data, newFilePath) {
    if (!fs.existsSync(path.dirname(newFilePath))) {
        fs.mkdirSync(path.dirname(newFilePath), { recursive: true });
    }
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, newFilePath);
}

async function main() {
    try {
        // Replace with the actual path to your Excel file
        const excelFilePath = '/Users/lilhesham/Downloads/AbdoAfter_win32_sorted.xlsx';
        const os = OS.WINDOWS;
        const sortedData = sortFileData(readExcelData(excelFilePath), os);

        // Replace with the desired path and filename for the new Excel file
        const outputDir = './output';
        const originalFileName = path.basename(excelFilePath);
        const outputFileName = originalFileName.replace('.xlsx', `_${os.toLowerCase()}_sorted.xlsx`);
        const newExcelFilePath = path.join(outputDir, outputFileName);
        writeExcelData(sortedData, newExcelFilePath);

        console.log('Excel file sorted and saved successfully!');
    } catch (error) {
        console.error('Error:', error);
    }
}

main();
