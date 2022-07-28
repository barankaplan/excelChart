import * as fs from "fs/promises";
import JSZip from "jszip";
import { Parser, Builder } from "xml2js";

var zip = new JSZip();
var parser = new Parser({ explicitArray: false });
var builder = new Builder();

function cloneDeep(obj) {
  const myDeepCopy = JSON.parse(JSON.stringify(obj));
  return myDeepCopy;
}

// next step: Do the same for a grouped bar chart

async function groupedBarChart(groupedBarChartData) {
  // read template
  const templatePath = "template/GroupedBarChart.xlsx";
  const template = await fs.readFile(templatePath);
  // unzip
  const content = await zip.loadAsync(template);
 // console.log("kaplancont")
  //console.log(content);
  const lengthOfData = groupedBarChartData.length;

  // three files must be changed: sheet2.xml, chart1.xml and sharedStrings.xml

  const dataAsString_sheet2 = await content
    .file(["xl/worksheets/sheet2.xml"])
    .async("string");
  const dataAsString_chart1 = await content
    .file(["xl/charts/chart1.xml"])
    .async("string");
  //console.log(dataAsString_chart1)
  const dataAsString_sharedStrings = await content
    .file(["xl/sharedStrings.xml"])
    .async("string");

  const dataParsed_sheet2 = await parser.parseStringPromise(
    dataAsString_sheet2
  );
  const dataParsed_chart1 = await parser.parseStringPromise(
    dataAsString_chart1
  );
  const dataParsed_sharedStrings = await parser.parseStringPromise(
    dataAsString_sharedStrings
  );

  // change sheet2.xml
  const headerRow = dataParsed_sheet2.worksheet.sheetData.row[0];
  headerRow.c[1].v = lengthOfData-2;

  const dataRow = dataParsed_sheet2.worksheet.sheetData.row[2];
  const newRows = [];

  newRows.push(headerRow);
  for (let i = 1; i < lengthOfData; i++) {
    const row = cloneDeep(dataRow);
    const rowNum = i + 1;
    row.c[0].v = i;
    row.c[1].v = groupedBarChartData[i][1];
    row.c[2].v = groupedBarChartData[i][2];

    row["$"].r = rowNum;
    row.c[0]["$"].r = "A" + rowNum;
    row.c[1]["$"].r = "B" + rowNum;
    row.c[2]["$"].r = "C" + rowNum;
    newRows.push(row);
    dataParsed_sheet2.worksheet.sheetData.row[i] = row;
  }

  // change sharedStrings.xml

  for (let i = lengthOfData; i < 3; i--) {
    dataParsed_sharedStrings.sst.si[i] = { t: groupedBarChartData[i][0] };
     console.log(dataParsed_sharedStrings.sst.si[i])
  }

  dataParsed_sharedStrings.sst.si[lengthOfData] = { t: groupedBarChartData[0][1] };
/*
const groupedBarChartData = [
  ["Häufigkeit", "A", "B"],
  ["Berlin", "3", "7"],
  ["München", "5", "7"],
  ["Hamburg", "4", "7"],
  ["Bremen", "7", "7"],
];
 */
  // change chart1.xml

  dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][
    "c:ser"]["c:cat"]["c:strRef"]["c:f"] = "Data!$A$2:$A$" + lengthOfData;

  dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][
    "c:ser"
  ]["c:val"]["c:numRef"]["c:f"] = "Data!$B$2:$B$" + lengthOfData;



  var xmlText_sheet2 = builder.buildObject(dataParsed_sheet2);
  var xmlText_sharedStrings = builder.buildObject(dataParsed_sharedStrings);
  var xmlText_chart1 = builder.buildObject(dataParsed_chart1);

  content.file(["xl/worksheets/sheet2.xml"], xmlText_sheet2);
  content.file(["xl/sharedStrings.xml"], xmlText_sharedStrings);
  content.file(["xl/charts/chart1.xml"], xmlText_chart1);

  const endFile = await zip.generateAsync({ type: "uint8array" });
  await fs.writeFile("template/GroupedBarChartKaplan.xlsx", endFile);
}

const groupedBarChartData = [
  ["Häufigkeit", "A", "B"],
  ["Berlin", "3", "7"],
  ["München", "5", "7"],
  ["Hamburg", "4", "7"],
  ["Bremen", "7", "7"],
];
groupedBarChart(groupedBarChartData);





