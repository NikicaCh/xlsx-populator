const Express = require("express");
const XlsxPopulate = require('xlsx-populate');
const bodyParser = require("body-parser");


let file = "";


let app = Express();

let port = process.env.PORT || 8000;


app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json({limit: '50mb', extended: true}))

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
    console.log(data[0])
    XlsxPopulate.fromFileAsync("./NTK-MAKS dnevni izvestaj.xlsx")
    .then(workbook => {
        data[0].map((obj, index) => {
            workbook.sheet("Sheet2").cell(`A${index+1}`).value(obj.Art);
            workbook.sheet("Sheet2").cell(`B${index+1}`).value(obj.Bolla);
            workbook.sheet("Sheet2").cell(`C${index+1}`).value(obj.Quantity);
        })
        data[1].map((obj, index) => {
            workbook.sheet("Sheet2").cell(`E${index+1}`).value(obj.Art);
            workbook.sheet("Sheet2").cell(`F${index+1}`).value(obj.Bolla);
            workbook.sheet("Sheet2").cell(`G${index+1}`).value(obj.Quantity);
        })
        workbook.outputAsync("base64")
        .then((data) => {
            file = data;
        })
    })
})

app.get("/download", (req, res) => {
    res.send(file)
})





app.listen(port, () => {
    console.log("Listening on port 8000")
})