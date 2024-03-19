import fs from 'node:fs';
import xlsx from 'node-xlsx';
import * as csv from 'csv';

const args = process.argv.slice(2);
const [command, ...rest] = args;

const commands = {
    // Converts an excel file to a csv file
    'convert-csv': (filePath) => {
        // Check if file exists
        if (!fs.existsSync(filePath)) throw new Error('File does not exist');
        // Check if we are parsing an excel file
        if (!filePath.endsWith('.xlsx')) throw new Error('File is not an excel file');
        // Parse the excel file
        const sheets = xlsx.parse(filePath);
        const ws = fs.createWriteStream(filePath.replace('.xlsx', '.csv'));
        // Convert the data to csv
        csv.stringify(sheets[0].data, (err, data) => {
            if (err) throw new Error(err);
            // Remove existing double quotes
            data = data.replace(/"/g, '');
            // For each line add a double quote at the beginning and end but not at the end of the file
            data = data.split('\n').map((line) => `"${line}"`).join('\n');
            // Add a double quote at the beginning and end of each comma unless it is next to a letter
            data = data.replace(/,(?=[a-zA-Z])/g, '","');
            // Read each line and if it equals "" remove it
            data = data.split('\n').filter((line) => line !== '""').join('\n');
            ws.write(data);
        });

        const file = fs.readFileSync(filePath, 'utf8');

        ws.on('finish', () => {
            fs.writeFileSync(filePath, file);
        });

        ws.on('error', (err) => {
            throw new Error(err);
        });
    },
    'filter': (filePath, headers) => {
        // Check if file exists
        if (!fs.existsSync(filePath, 'utf8')) throw new Error('File does not exist');
        // Check if headers are provided
        if (!headers) throw new Error('Headers are required');
        // Check if we are parsing a csv file
        if (!filePath.endsWith('.csv')) throw new Error('File is not a csv file');
        // Read the file
        const file = fs.readFileSync(filePath, 'utf8');
        // Get Headers
        let _headers = file.split('\n')[0].split(',');
        // Convert headers to array
        headers = headers.split(',');
        // Check if headers are valid
        const invalidHeaders = headers.filter((header) => !_headers.includes(`"${header}"`));
        if (invalidHeaders.length) throw new Error(`Invalid headers: ${invalidHeaders.join(', ')}`);

        // Get rows
        const rows = file.split('\n').slice(1);
        // Filter the rows by the specified headers
        const _rows = rows.map((row) => {
            const _row = row.split(',');
            return headers.map((header) => {
                const index = _headers.indexOf(`"${header}"`);
                return _row[index];
            });
        });

        // Only use the headers that were specified and don't write duplicates
        fs.writeFileSync(filePath.replace('.csv', '.filtered.csv'), `"${headers.join('","')}"\n${_rows.join('\n')}`);
        // Remove duplicate rows from filteredFile
        const filteredFileData = fs.readFileSync(filePath.replace('.csv', '.filtered.csv'), 'utf8');
        const filteredFileRows = filteredFileData.split('\n').slice(1);
        const filteredFileRowsSet = new Set(filteredFileRows);
        const filteredFileRowsArray = Array.from(filteredFileRowsSet);
        fs.writeFileSync(filePath.replace('.csv', '.filtered.csv'), `"${headers.join('","')}"\n${filteredFileRowsArray.join('\n')}`);
    },
    'convert-excel': (filePath) => {
        // Check if file exists
        if (!fs.existsSync(filePath)) throw new Error('File does not exist');
        // Check if we are parsing a csv
        if (!filePath.endsWith('.csv')) throw new Error('File is not a csv file');
        // Read the file
        const ws = fs.createWriteStream (filePath.replace('.csv', '.xlsx'));
        const file = fs.readFileSync(filePath, 'utf8');
        let data = file.split('\n').map((row) => row.split(','));
        // Remove double quotes in a row from the data
        data = data.map((row) => row.map((cell) => cell.replace(/"/g, '')));
        const buffer = xlsx.build([{name: 'Sheet1', data}]);
        ws.write(buffer);
        ws.end();
    },
    'sort': (filePath) => {
        // Check if file exists
        if (!fs.existsSync(filePath)) throw new Error('File does not exist');
        // Check if we are parsing a csv
        if (!filePath.endsWith('.csv')) throw new Error('File is not a csv file');
        // Read the file
        const file = fs.readFileSync(filePath, 'utf8');
        // Get headers
        const headers = file.split('\n')[0].split(',');
        // Get rows
        const rows = file.split('\n').slice(1);
        // Sort the rows
        const sortedRows = rows.sort();
        // Write the file
        fs.writeFileSync(filePath.replace('.csv', '.sorted.csv'), `${headers.join(',')}\n${sortedRows.join('\n')}`);
    }
};

if (commands[command]) {
    commands[command](rest[0], ...rest.slice(1) || []);
} else {
    console.log('Command not found');
}