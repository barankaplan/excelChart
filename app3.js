import * as fs from "fs/promises";
import JSZip from "jszip";
import {Parser, Builder} from "xml2js";

var zip = new JSZip();
var parser = new Parser({explicitArray: false});
var builder = new Builder();

function cloneDeep(obj) {
    const myDeepCopy = JSON.parse(JSON.stringify(obj));
    return myDeepCopy;
}

async function simpleBarChart(dataToPlot) {
    // read template
    const templatePath = "template/GroupedBarChart.xlsx";
    const template = await fs.readFile(templatePath);
    // unzip
    const content = await zip.loadAsync(template);
    const lengthOfData = dataToPlot.length;

    // three files must be changed: sheet2.xml, chart1.xml and sharedStrings.xml

    const dataAsString_sheet2 = await content
        .file(["xl/worksheets/sheet2.xml"])
        .async("string");
    const dataAsString_chart1 = await content
        .file(["xl/charts/chart1.xml"])
        .async("string");
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

    const header = dataParsed_sheet2.worksheet.sheetData.row[0];

    header.c[0].v = dataToPlot[0][0];
    header.c[1].v = dataToPlot[0][1];
    header.c[2].v = dataToPlot[0][2];


    const secondRow = dataParsed_sheet2.worksheet.sheetData.row[1];

    const newRows = [];


    newRows.push(header);

    for (let i = 1; i < lengthOfData; i++) {
        const row = cloneDeep(secondRow);
        const rowNum = i + 1;
        row.c[0].v = dataToPlot[i][0];
        row.c[1].v = dataToPlot[i][1];
        row.c[2].v = dataToPlot[i][2];


        row["$"].r = rowNum;
        row.c[0]["$"].r = "A" + rowNum;
        row.c[1]["$"].r = "B" + rowNum;
        row.c[2]["$"].r = "C" + rowNum;
        newRows.push(row);
        dataParsed_sheet2.worksheet.sheetData.row[i] = row;
    }


    // change sharedStrings.xml
    /*
        for (let i =0 ; i < lengthOfData-2; i++) {
            dataParsed_sharedStrings.sst.si[i] = {t: dataToPlot[i+1][0]};
        }
            //header
            dataParsed_sharedStrings.sst.si[lengthOfData] = {t: dataToPlot[0][0]};

     */
    const sharedStrings = [];



    for (let i = 1; i < lengthOfData; i++) {

        sharedStrings.push({t: dataToPlot[i][0]});

    }
    for (let i = 0; i < dataToPlot[0].length; i++) {

        sharedStrings.push({t: dataToPlot[0][i]});

    }

    console.log(sharedStrings);


    console.log(dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"][0]["c:cat"]["c:strRef"]["c:f"]);
    console.log(dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"][0]["c:val"]["c:numRef"]["c:f"]);
    console.log(dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"][1]["c:val"]["c:numRef"]["c:f"]);

    // change chart1.xml


    dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][
        "c:ser"
        ][0]["c:cat"]["c:strRef"]["c:f"] = "Data!$A$2:$A$" + lengthOfData;

    dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][
        "c:ser"
        ][0]["c:val"]["c:numRef"]["c:f"] = "Data!$B$2:$B$" + lengthOfData;

    dataParsed_chart1["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][
        "c:ser"
        ][1]["c:val"]["c:numRef"]["c:f"] = "Data!$C$2:$C$" + lengthOfData;




    var xmlText_sheet2 = builder.buildObject(dataParsed_sheet2);
    var xmlText_sharedStrings = builder.buildObject(dataParsed_sharedStrings);
    var xmlText_chart1 = builder.buildObject(dataParsed_chart1);

    content.file(["xl/worksheets/sheet2.xml"], xmlText_sheet2);
    content.file(["xl/sharedStrings.xml"], xmlText_sharedStrings);
    content.file(["xl/charts/chart1.xml"], xmlText_chart1);

    const endFile = await zip.generateAsync({type: "uint8array"});
    await fs.writeFile("BaranGroupedBarChart.xlsx", endFile);


}

const simpleBarChartData = [
    ["A", "B", "C"],
    ["Berlin", "0", "1"],
    ["München", "2", "3"],
    ["Hamburg", "4", "5"],
    ["Bremen", "6", "7"],
    ["Bielefeld", "8", "9"],
    ["Istanbul", "10", "11"]
];
simpleBarChart(simpleBarChartData);

// next step: Do the same for a grouped bar chart


/*

function groupedBarChart(groupedBarChartData) {
  // to do Baran
}

const groupedBarChartData = [
  ["Häufigkeit", "A", "B", "c"],
  ["Berlin", "7", "7", "7"],
  ["München", "5", "7", "5"],
  ["Hamburg", "4", "7", "5"],
  ["Bremen", "7", "7", "5"],
];
groupedBarChart(groupedBarChartData);

 */


