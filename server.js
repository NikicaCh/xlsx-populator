const Express = require("express");
const XlsxPopulate = require('xlsx-populate');
const bodyParser = require("body-parser");


let app = Express();

let port = process.env.PORT || 8000;

app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json())

app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers",  "Content-Type");
    next();
});
app.get("/", (req, res) => {
    res.send("HELLO")
})

app.post("/", (req, res) => {
    res.set('Content-Type', 'application/json');
    let data = req.body.data;
    res.status(201);
    res.json();
    XlsxPopulate.fromFileAsync("../NTK-MAKS dnevni izvestaj.xlsx")
    .then(workbook => {
        data.map((obj, index) => {
            workbook.sheet("Sheet2").cell(`A${index+1}`).value(obj.name);
            workbook.sheet("Sheet2").cell(`B${index+1}`).value(obj.id);
            workbook.sheet("Sheet2").cell(`C${index+1}`).value(obj.quantity);
        })
        return workbook.toFileAsync(`../NTK-MAKS dnevni izvestaj 04.09.2019.xlsx`);
    });
})



app.listen(port, () => {
    console.log("Listening on port 8000")
})