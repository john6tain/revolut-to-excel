const csv = require("csvtojson");
const {google} = require('googleapis');
const {readFile} = require('fs').promises;

const csvFilePath = './data.csv'
/*
*  TODO: read all older sheets for check for more values in categories
* */

const categories = {
    'ПАЗАР': {columns: 'A3:B14', values: ['Tabako Ilevan Ood', 'KAM MARKET']},
    'РЕСТОРАНТИ': {columns: 'C3:D14', values: ['Takeaway.com', 'ReZZo']},
    'ГОРИВО': {columns: 'E3:F14', values: ['Evpoint App']},
    'ЗАБАВЛЕНИЯ': {columns: 'G3:H14', values: ['Pepe End Sil Eood']},
    'ЛЕКАРСВА': {columns: 'I3:J14', values: ['Nadezhda Hospital', 'Angelina Orfey Bakalova', 'dm drogerie']},
    'СМЕТКИ': {columns: 'K3:L14', values: ['Vivacom', 'Netflix']},
    'ЗА ДОМА': {columns: 'M3:N14', values: []},
    'ПОДАРАЦИ': {columns: 'O3:P14', values: ['Amazon']},
    'UNIDENTIFIED': {columns: 'Q3:R1000', values: []},
}
categories['UNIDENTIFIED'].values = Object.keys(categories).map(key => categories[key].values).reduce((accumulator, currentArray) => {
    return accumulator.concat(currentArray);
}, []);


let jsonData;

function prepareCategories() {
    categories['ПАЗАР'].rows = jsonData.filter(row => categories['ПАЗАР'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['РЕСТОРАНТИ'].rows = jsonData.filter(row => categories['РЕСТОРАНТИ'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['ГОРИВО'].rows = jsonData.filter(row => categories['ГОРИВО'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['ЗАБАВЛЕНИЯ'].rows = jsonData.filter(row => categories['ЗАБАВЛЕНИЯ'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['ЛЕКАРСВА'].rows = jsonData.filter(row => categories['ЛЕКАРСВА'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['СМЕТКИ'].rows = jsonData.filter(row => categories['СМЕТКИ'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['ЗА ДОМА'].rows = jsonData.filter(row => categories['ЗА ДОМА'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['ПОДАРАЦИ'].rows = jsonData.filter(row => categories['ПОДАРАЦИ'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));
    categories['UNIDENTIFIED'].rows = jsonData.filter(row => !categories['UNIDENTIFIED'].values.some(el => el.toLowerCase() === row[0].toLowerCase()));

    Object.keys(categories).forEach(key => {
        const items = {};
        categories[key].rows.forEach(([item, value]) => {
            if (items[item]) {
                items[item] += value;
            } else {
                items[item] = value;
            }
        });
        categories[key].rows = Object.entries(items)
    });
}

async function extractCSV() {
    jsonData = await csv().fromFile(csvFilePath);
    jsonData = jsonData.filter(el => el.Amount < 0).map(el => ([el['Description'], Math.round(Math.abs(el['Amount']))]));
    prepareCategories();
}


async function loadCredentials() {
    const content = await readFile('credentials.json');
    return JSON.parse(content);
}

async function authenticate() {
    const credentials = await loadCredentials();

    const auth = new google.auth.GoogleAuth({
        credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    return auth;
}

async function insertData(auth) {
    // console.log(jsonData) ;
    // process.exit(0);

    let index = 0;
    for (const key of Object.keys(categories)) {


        const sheets = google.sheets({version: 'v4', auth});
        const spreadsheetId = '1XgQ-MJFE9s9V3PAJNBPtOLL2sRMYKb4xIeXqlUrW5sM';

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `Sheet1!${categories[key].columns}`,
        });

        const rows = response.data.values || [];
        const insertFrom = rows.length + 3;
        categories[key].columns = categories[key].columns.replace('3', `${insertFrom}`)
        const range = `Sheet1!${categories[key].columns}`
        const resource = {
            values: categories[key].rows
        };
        // console.log('start',insertFrom,'end',categories[key].rows.length +1);

        const requests = [
            {
                repeatCell: {
                    range: {
                        sheetId: 0,
                        startRowIndex: insertFrom - 1,
                        endRowIndex: insertFrom + categories[key].rows.length - 1,
                        startColumnIndex: index * 2,
                        endColumnIndex: (index * 2) + 2
                    },
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: {
                                red: 1,
                                green: 1,
                                blue: 0.6,
                            },
                        },
                    },
                    fields: 'userEnteredFormat.backgroundColor',
                },
            },
        ];


        try {
            const response = await sheets.spreadsheets.values.update({
                spreadsheetId, range, valueInputOption: 'RAW', resource,
            });
            const update = await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests,
                },
            });
            console.log(`${response.data.updatedCells} cells updated.`);
        } catch (err) {
            console.error('Error inserting data: ', err);
            // console.error('Error inserting data: ', JSON.stringify(err));
        }
        index++;
    }
}


extractCSV().then(data => authenticate().then(insertData).catch(console.error)).catch(console.error);
