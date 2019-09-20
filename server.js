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


app.post("/report", (req, res, next) => {
    res.set('Content-Type', 'application/json');
    let data = req.body.data;
    res.status(201);
    res.json();
    XlsxPopulate.fromFileAsync("./NTK-MAKS dnevni izvestaj.xlsx")
    .then(workbook => {
        data[0].map((obj, index) => {
            workbook.sheet("Sheet2").cell(`A${index+1}`).value(obj.Art);
            workbook.sheet("Sheet2").cell(`B${index+1}`).value(obj.Bolla);
            workbook.sheet("Sheet2").cell(`C${index+1}`).value(parseInt(obj.Quantity));
        })
        data[1].map((obj, index) => {
            workbook.sheet("Sheet2").cell(`E${index+1}`).value(obj.Art);
            workbook.sheet("Sheet2").cell(`F${index+1}`).value(obj.Bolla);
            workbook.sheet("Sheet2").cell(`G${index+1}`).value(parseInt(obj.Quantity));
        })
        workbook.outputAsync("base64")
        .then((data) => {
            file = data;
        })
    })
    next()
})

const groupBy = key => array =>
  array.reduce((objectsByKeyValue, obj) => {
    const value = obj[key];
    objectsByKeyValue[value] = (objectsByKeyValue[value] || []).concat(obj);
    return objectsByKeyValue;
  }, {});

let exportFile = "";


app.post("/export", (req, res) => {
    res.set('Content-Type', 'application/json');
    let data = req.body.data;
    let a=0, b=0, c=0, d=0; //populated spaces in 4 columns
    res.status(201);
    res.json();
    let groupByName = groupBy("Art");
    let obj = groupByName(data)
    XlsxPopulate.fromFileAsync("./packing lista.xlsx")
    .then( workbook => {
        let totalA = 0, totalB = 0, totalC = 0, totalD = 0;
        Object.keys(obj).map((key, index) => {
            let objLength = obj[key].length
            if(objLength +1 <= (21-a)) {
                let total = 0;
                let totalIndex = 0;
                Object.keys(obj[key]).map((_key, _index) => {
                    //code goes here
                    let index = a + 6 + _index;
                    totalIndex = index;
                    total += parseInt(obj[key][_key].Quantity);
                    workbook.sheet("Sheet1").cell(`B${index}`).value(obj[key][_key].Art)
                    workbook.sheet("Sheet1").cell(`C${index}`).value(obj[key][_key].Bolla)
                    workbook.sheet("Sheet1").cell(`D${index}`).value(obj[key][_key].Quantity)
                    workbook.sheet("Sheet1").cell(`E${index}`).value(obj[key][_key].ImportNo)
                })
                workbook.sheet("Sheet1").range(`A${totalIndex +1}:E${totalIndex +1}`).style("borderStyle", "medium")
                workbook.sheet("Sheet1").cell(`B${totalIndex +1}`).value("VKUPNO").style("bold", true)
                workbook.sheet("Sheet1").cell(`D${totalIndex +1}`).value(total)
                totalA += total;
                workbook.sheet("Sheet1").cell("D27").value(totalA)
                a += objLength + 1;
            } else if(objLength + 1 <= (21 - b)) {
                let total = 0;
                let totalIndex = 0;
                Object.keys(obj[key]).map((_key, _index) => {
                    //code goes here
                    let index = b + 6 +_index;
                    totalIndex = index;
                    total += parseInt(obj[key][_key].Quantity);
                    workbook.sheet("Sheet1").cell(`G${index}`).value(obj[key][_key].Art)
                    workbook.sheet("Sheet1").cell(`H${index}`).value(obj[key][_key].Bolla)
                    workbook.sheet("Sheet1").cell(`I${index}`).value(obj[key][_key].Quantity)
                    workbook.sheet("Sheet1").cell(`J${index}`).value(obj[key][_key].ImportNo)
                })
                workbook.sheet("Sheet1").range(`F${totalIndex +1}:J${totalIndex +1}`).style("borderStyle", "medium")
                workbook.sheet("Sheet1").cell(`G${totalIndex +1}`).value("VKUPNO").style("bold", true)
                workbook.sheet("Sheet1").cell(`I${totalIndex +1}`).value(total)
                totalB += total;
                workbook.sheet("Sheet1").cell("I27").value(totalB)
                b += objLength + 1;
            } else if(objLength + 1 <= (21 - c)) {
                let total = 0;
                let totalIndex = 0;
                total += parseInt(obj[key][_key].Quantity);
                Object.keys(obj[key]).map((_key, _index) => {
                    //code goes here
                    let index = c + 6 + _index;
                    totalIndex = index;
                    workbook.sheet("Sheet1").cell(`L${index}`).value(obj[key][_key].Art)
                    workbook.sheet("Sheet1").cell(`M${index}`).value(obj[key][_key].Bolla)
                    workbook.sheet("Sheet1").cell(`N${index}`).value(obj[key][_key].Quantity)
                    workbook.sheet("Sheet1").cell(`O${index}`).value(obj[key][_key].ImportNo)
                })
                workbook.sheet("Sheet1").range(`K${totalIndex +1}:O${totalIndex +1}`).style("borderStyle", "medium")
                workbook.sheet("Sheet1").cell(`L${totalIndex +1}`).value("VKUPNO").style("bold", true)
                workbook.sheet("Sheet1").cell(`N${totalIndex +1}`).value(total)
                totalC += total;
                workbook.sheet("Sheet1").cell("N27").value(totalC)
                c += objLength + 1;
            } else if(objLength + 1 <= (21 - d)) {
                let total = 0;
                let totalIndex = 0;
                total += parseInt(obj[key][_key].Quantity);
                Object.keys(obj[key]).map((_key, _index) => {
                    //code goes here
                    let index = d + 6 + _index;
                    totalIndex = index;
                    workbook.sheet("Sheet1").cell(`Q${index}`).value(obj[key][_key].Art)
                    workbook.sheet("Sheet1").cell(`R${index}`).value(obj[key][_key].Bolla)
                    workbook.sheet("Sheet1").cell(`S${index}`).value(obj[key][_key].Quantity)
                    workbook.sheet("Sheet1").cell(`T${index}`).value(obj[key][_key].ImportNo)
                })
                workbook.sheet("Sheet1").range(`P${totalIndex +1}:T${totalIndex +1}`).style("borderStyle", "medium")
                workbook.sheet("Sheet1").cell(`Q${totalIndex +1}`).value("VKUPNO").style("bold", true)
                workbook.sheet("Sheet1").cell(`S${totalIndex +1}`).value(total)
                totalD += total;
                workbook.sheet("Sheet1").cell("S27").value(totalD)
                d += objLength + 1;
            }
        })
        workbook.outputAsync("base64")
            .then((data) => {
                exportFile = data;
            })
        
    })
    
})

app.post('/wakeup', (req, res) => {
    res.status(201)
})

app.get("/download", (req, res) => {
    res.send(file)
})

app.get("/export", (req, res) => {
    res.send(exportFile)
})





app.listen(port, () => {
    console.log("Listening on port 8000")
})