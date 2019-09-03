const Express = require("express");
const XlsxPopulate = require('xlsx-populate');

const Port = process.env.Port || 3000



let app = Express();


app.get("/", (req, res) => {
    XlsxPopulate.fromFileAsync("../Book1.xlsx")
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("G1").value("NIKICA MAKSIMOVSKI");
 
        // Write to file.
        return workbook.toFileAsync("../Book1.xlsx");
    });
    res.send("HELLO")
})



app.listen(Port, () => {
    console.log(`Listening on port ${Port}`)
})