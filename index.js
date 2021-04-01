const mysql = require('mysql');
const exceljs = require('exceljs');

const workbook = new exceljs.Workbook();
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

async function writeExcel() {
    await con.connect(async (err) => {
        if (err) {
            throw err;
        }
    
        await con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'standard' AND locale = 'domestic'", async function (err, result, fields) {
            if (err) {
                throw err;
            }

            const keys = Object.keys(result[0]);
            const innerFunc = () => {
                worksheet.columns = keys.map(header => {
                    return {
                        header: header,
                        key: header
                    }
                })
            }
            
            await innerFunc();
            await worksheet.addRows(result);
        })
    
        await con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'expedited' AND locale = 'domestic'", async function (err, result, fields) {
            if (err) {
                throw err;
            }
    
            const keys = Object.keys(result[0]);
            const innerFunc = () => {
                worksheet1.columns = keys.map(header => {
                    return {
                        header: header,
                        key: header
                    }
                })
            }
            
            await innerFunc();
            await worksheet1.addRows(result);
        })
    
        await con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'nextDay' AND locale = 'domestic'", async function (err, result, fields) {
            if (err) {
                throw err;
            }
    
            const keys = Object.keys(result[0]);
            const innerFunc = () => {
                worksheet2.columns = keys.map(header => {
                    return {
                        header: header,
                        key: header
                    }
                })
            }
            
            await innerFunc();
            await worksheet2.addRows(result);
        })
    
        await con.query("SELECT * from rates WHERE client_id = 1240 AND shipping_speed = 'intlEconomy' AND locale = 'international'", async function (err, result, fields) {
            if (err) {
                throw err;
            }
    
            const keys = Object.keys(result[0]);
            const innerFunc = () => {
                worksheet3.columns = keys.map(header => {
                    return {
                        header: header,
                        key: header
                    }
                })
            }
            
            await innerFunc();
            await worksheet3.addRows(result);
        })
    })

    setTimeout(async () => {
        await workbook.xlsx.writeFile('output.xlsx');
    }, 10000)
    
}

writeExcel();