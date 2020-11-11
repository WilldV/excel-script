const express = require("express");
const fs = require("fs");
const path = require("path");
const excelToJson = require("convert-excel-to-json");
const app = express();
var json2xls = require('json2xls');

app.set("PORT", process.env.PORT || 3000);

app.get("/", async (req, res) => {
  try {
    let result = excelToJson({
      sourceFile: path.join(__dirname, "DATA_FINAL.xls"),
    });
    const r = result.Sheet1
    const a = r[0]
    const b = Object.keys(r[1])
    r.splice(0,1)
    const values = Object.values(a)
    const data = r;
    const filteredData = []
    data.forEach(e => {
      let flag = true
      if(e.A < 1 || e.A > 4) flag = false
      if(e.B < 1 || e.B > 2) flag = false
      if(e.D < 1 || e.D > 2) flag = false
      if(e.E < 1 || e.E > 6) flag = false
      if(e.F < 1 || e.F > 4) flag = false
      if(e.G < 0 || e.G > 70) flag = false
      if(e.H < 0 || e.H > 70) flag = false
      if(e.J < 1 || e.J > 5) flag = false
      if(e.K < 1 || e.K > 5) flag = false
      if(e.L < 1 || e.L > 5) flag = false
      if(e.M < 1 || e.M > 5) flag = false
      if(e.N < 0 || e.N > 6) flag = false
      if(e.O < 0 || e.O > 6) flag = false
      if(e.P < 0 || e.P > 6) flag = false
      if(e.Q < 0 || e.Q > 6) flag = false
      if(e.R < 0 || e.R > 6) flag = false
      if(e.S < 0 || e.S > 6) flag = false
      if(e.T < 0 || e.T > 6) flag = false
      if(e.U < 0 || e.U > 6) flag = false
      if(e.V < 0 || e.V > 6) flag = false
      if(e.W < 0 || e.W > 6) flag = false
      if(e.X < 0 || e.X > 6) flag = false
      if(e.Y < 0 || e.Y > 6) flag = false
      if(e.Z < 0 || e.Z > 6) flag = false
      if(e.AA < 0 || e.AA > 6) flag = false
      if(e.AB < 0 || e.AB > 6) flag = false
      if(e.AC < 0 || e.AC > 6) flag = false
      if(e.AD < 0 || e.AD > 6) flag = false
      if(e.AE < 0 || e.AE > 6) flag = false
      if(e.AF < 0 || e.AF > 6) flag = false
      if(e.AG < 0 || e.AG > 6) flag = false
      if(e.AH < 0 || e.AH > 6) flag = false
      if(e.AI < 0 || e.AI > 6) flag = false

      if(flag){
        let newObject = {}
        values.forEach((k, index) => {
          newObject[k] = e[b[index]]
        });
        filteredData.push(newObject)
      } 
    });
    var xls = json2xls(filteredData);
    fs.writeFileSync('data.xlsx', xls, 'binary');
    return res.send(filteredData);
  } catch (error) {
    console.log(error);
  }
});

app.listen(app.get("PORT"), () => {
  console.log(`Server on PORT: ${app.get("PORT")}`);
});
