var Excel = require('exceljs');

const filename = './resources/test.xlsx';

const data = [
    {
        firstName: 'Yurii',
        lastName: 'Khm',
    },
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
    {
        firstName: 'Karolina',
        lastName: 'Nick',
    },    
];

console.log('Start creating excel file ================');

const workbook = new Excel.Workbook();
const sheet = workbook.addWorksheet('Mapping', {
    properties: {
        tabColor: { argb: 'FF005500' },
    }
});

// for (let idx = 0; idx < data.length; idx++) {
//     console.log(data[idx].firstName + ' ' + data[idx].lastName);
// }

// data.map(user => {
//     console.log(user.firstName + ' ' + user.lastName);
// });

// data.map(({ firstName, lastName }) =>
//     console.log(`${firstName} ${lastName}`)
// );
    
sheet.columns = [
    { header: 'First Name', key: 'firstName', width: 15 },
    { header: 'Last Name', key: 'lastName', width: 32 },
    { header: 'Value', key: 'asdf', width: 10 }
];

const adjustedData = data.map(user => ({
    ...user,
    asdf: Math.round(Math.random(1000) * 1000),
}));

sheet.addRows(adjustedData);

const countUsers = adjustedData.length;

sheet.getCell(`C${2 + countUsers}`).value =
    { formula: `SUM(C2:C${1 + countUsers})` };

workbook.xlsx.writeFile(filename).then(done => {
    console.log('Create file success!');
});