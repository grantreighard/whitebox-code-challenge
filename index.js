const mysql = require('mysql');
const excel = require('excel4node');

const workbook = new excel.Workbook();
const worksheet = workbook.addWorksheet('Domestic Standard Rates');
const worksheet1 = workbook.addWorksheet('Domestic Expedited Rates');
const worksheet2 = workbook.addWorksheet('Domestic Next Day Rates');
const worksheet3 = workbook.addWorksheet('International Economy Rates');

const con = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: 'hello',
    database: 'data'
})

const style = workbook.createStyle({
    font: {
        color: '#000000',
        size: 12
    }
})

con.connect((err) => {
    if (err) {
        throw err;
    }

    con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'standard' AND locale = 'domestic'", function (err, result, fields) {
        if (err) {
            throw err;
        }

        for(let i = 0; i < result.length; i++) {
            for (let j = 0; j < Object.keys(result[i]).length; j++) {
                const data = result[i][Object.keys(result[i])[j]];
                typeof data === 'number' ? worksheet.cell(i, j).number(data).style(style) : worksheet.cell(i, j).string(data).style(style);
            }
        }
    })

    con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'expedited' AND locale = 'domestic'", function (err, result, fields) {
        if (err) {
            throw err;
        }

        for(let i = 0; i < result.length; i++) {
            for (let j = 0; j < Object.keys(result[i]).length; j++) {
                const data = result[i][Object.keys(result[i])[j]];
                typeof data === 'number' ? worksheet1.cell(i, j).number(data).style(style) : worksheet.cell(i, j).string(data).style(style);
            }
        }
    })

    con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'next_day' AND locale = 'domestic'", function (err, result, fields) {
        if (err) {
            throw err;
        }

        for(let i = 0; i < result.length; i++) {
            for (let j = 0; j < Object.keys(result[i]).length; j++) {
                const data = result[i][Object.keys(result[i])[j]];
                typeof data === 'number' ? worksheet2.cell(i, j).number(data).style(style) : worksheet.cell(i, j).string(data).style(style);
            }
        }
    })

    con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'economy' AND locale = 'international'", function (err, result, fields) {
        if (err) {
            throw err;
        }

        for(let i = 0; i < result.length; i++) {
            for (let j = 0; j < Object.keys(result[i]).length; j++) {
                const data = result[i][Object.keys(result[i])[j]];
                typeof data === 'number' ? worksheet3.cell(i, j).number(data).style(style) : worksheet.cell(i, j).string(data).style(style);
            }
        }
    })

    workbook.write('output.xlsx');
})