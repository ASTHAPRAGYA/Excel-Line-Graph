// ======================================================
// Solar Inverter Temperature Analyzer
// Complete script.js
// ======================================================


let workbook = null;

let workbookData = {};

let combinedData = [];

let tempChart = null;



const fileInput = document.getElementById("excelFile");

const selectionCard = document.getElementById("selectionCard");

const checkboxContainer = document.getElementById("checkboxContainer");

const generateBtn = document.getElementById("generateBtn");




// ======================================================
// Upload Excel File
// ======================================================

fileInput.addEventListener("change", function(event){


    const file = event.target.files[0];


    if(!file) return;



    const reader = new FileReader();



    reader.onload = function(e){


        const data = new Uint8Array(e.target.result);



        workbook = XLSX.read(data,{

            type:"array"

        });



        console.log(
            "Workbook Loaded:",
            workbook.SheetNames
        );



        detectSheets();


    };



    reader.readAsArrayBuffer(file);


});





// ======================================================
// Detect Worksheets
// ======================================================

function detectSheets(){


    workbookData = {};

    checkboxContainer.innerHTML = "";



    workbook.SheetNames.forEach(sheetName=>{


        console.log(
            "Checking:",
            sheetName
        );



        const sheet = workbook.Sheets[sheetName];



        const rows = XLSX.utils.sheet_to_json(sheet,{

            header:1,

            defval:""

        });



        let headerRow = -1;



        // Search first 15 rows

        for(
            let i=0;
            i<Math.min(rows.length,15);
            i++
        ){


            let headers = rows[i].map(h=>

                normalizeHeader(h)

            );



            if(

                headers.includes("TIMESTAMP") &&

                headers.includes("TEMP1") &&

                headers.includes("TEMP2")

            ){

                headerRow=i;

                break;

            }


        }



        if(headerRow===-1){


            console.log(
                sheetName,
                "Header not found"
            );


            return;


        }





        console.log(

            sheetName,

            "Header found at row",

            headerRow+1

        );





        let rawData =
        XLSX.utils.sheet_to_json(sheet,{

            range:headerRow,

            defval:""

        });





        let cleaned =
        cleanSheetData(rawData);





        if(cleaned.length>0){


            workbookData[sheetName]=cleaned;


        }



    });





    console.log(
        "Detected Inverters:",
        workbookData
    );





    if(
        Object.keys(workbookData).length===0
    ){

        alert(
            "No valid inverter worksheets found."
        );

        return;

    }



    buildSelectionTable();



    selectionCard.classList.remove(
        "hidden"
    );


}






// ======================================================
// Normalize Headers
// ======================================================

function normalizeHeader(value){


    return value
    .toString()
    .trim()
    .replace(/\s+/g,"")
    .toUpperCase();


}







// ======================================================
// Clean Sheet Data
// ======================================================

function cleanSheetData(data){


    return data.map(row=>{


        let newRow={};



        Object.keys(row).forEach(key=>{


            let clean =
            normalizeHeader(key);



            if(clean.includes("TIME")){


                newRow["Time Stamp"]
                =
                row[key];


            }



            else if(clean.includes("TEMP1")){


                newRow["TEMP1"]
                =
                Number(row[key]);


            }



            else if(clean.includes("TEMP2")){


                newRow["TEMP2"]
                =
                Number(row[key]);


            }



        });



        return newRow;


    })

    .filter(row=>

        row["Time Stamp"]

    );


}








// ======================================================
// Create Selection Table
// ======================================================

function buildSelectionTable(){


    checkboxContainer.innerHTML="";



    let table =
    document.createElement("table");



    table.innerHTML=`

    <tr>

        <th>Inverter</th>

        <th>TEMP1</th>

        <th>TEMP2</th>

    </tr>

    `;



    Object.keys(workbookData)
    .forEach(inv=>{


        let row =
        document.createElement("tr");



        row.innerHTML=`

        <td>
            ${inv}
        </td>


        <td>

            <input 
            type="checkbox"
            value="${inv}|TEMP1">

        </td>


        <td>

            <input
            type="checkbox"
            value="${inv}|TEMP2">

        </td>


        `;



        table.appendChild(row);



    });



    checkboxContainer.appendChild(table);



}







// ======================================================
// Get Selected Temperatures
// ======================================================

function getSelectedSeries(){


    let selected=[];



    document
    .querySelectorAll(
        "#checkboxContainer input:checked"
    )

    .forEach(box=>{


        let value =
        box.value.split("|");



        selected.push({

            inverter:value[0],

            temperature:value[1]

        });



    });



    return selected;


}









// ======================================================
// Generate Graph
// ======================================================

generateBtn.addEventListener(
"click",
function(){


    let selected =
    getSelectedSeries();



    if(selected.length===0){


        alert(
            "Select at least one temperature."
        );


        return;

    }



    createCombinedData();



    createChart(selected);



});









// ======================================================
// Create Combined Data
// ======================================================

function createCombinedData(){


    combinedData=[];



    let times = new Set();




    Object.keys(workbookData)
    .forEach(inv=>{


        workbookData[inv]
        .forEach(row=>{


            times.add(

               convertExcelDate(
    row["Time Stamp"]
)
.getTime()

            );


        });


    });





    let sortedTimes =
    Array.from(times)
    .sort(
        (a,b)=>a-b
    );





    sortedTimes.forEach(time=>{


        let row={};


        row["Time Stamp"]
        =
        new Date(time);



        Object.keys(workbookData)
        .forEach(inv=>{


            row[inv+" TEMP1"]="";

            row[inv+" TEMP2"]="";


        });



        combinedData.push(row);



    });






    Object.keys(workbookData)
    .forEach(inv=>{


        workbookData[inv]
        .forEach(data=>{


            let time =

           convertExcelDate(
    data["Time Stamp"]
)
.getTime();




            let index =
            sortedTimes.indexOf(time);




            if(index!==-1){


                combinedData[index]
                [inv+" TEMP1"]
                =
                data["TEMP1"];



                combinedData[index]
                [inv+" TEMP2"]
                =
                data["TEMP2"];


            }


        });



    });



    console.log(
        "Combined Data",
        combinedData
    );


}









// ======================================================
// Create Chart
// ======================================================

// ======================================================
// Create Advanced Temperature Chart
// ======================================================

function createChart(selected){


    const colors = [

        "#27A5AD",
        "#FF6384",
        "#36A2EB",
        "#FFCE56",
        "#9966FF",
        "#4BC0C0",
        "#FF9F40",
        "#8BC34A"

    ];



    let datasets=[];

    let allDates=[];



    selected.forEach((item,index)=>{


        let key =
        item.inverter +
        " " +
        item.temperature;



        let dataPoints =
        combinedData
        .map(row=>{


let rawTime = row["Time Stamp"];

let date;


if(typeof rawTime === "number"){

    date = XLSX.SSF.parse_date_code(rawTime);

    date = new Date(

        date.y,
        date.m-1,
        date.d,
        date.H,
        date.M,
        date.S

    );

}
else{

    date = new Date(rawTime);

}



            let value =
            row[key];



            if(

                isNaN(date.getTime()) ||

                value==="" ||

                value===undefined ||

                isNaN(value)

            ){

                return null;

            }



            allDates.push(date);



            return {

                x:date,

                y:Number(value)

            };


        })
        .filter(x=>x!==null);





        datasets.push({


            label:key,


            data:dataPoints,


            borderColor:
            colors[index % colors.length],



            borderWidth:2,


            pointRadius:0,


            tension:0.1,


            fill:false


        });



    });





    if(allDates.length===0){

        alert("No valid temperature data found");

        return;

    }






    // Actual Excel date range only

    let minDate =
    new Date(
        Math.min(
            ...allDates
        )
    );



    let maxDate =
    new Date(
        Math.max(
            ...allDates
        )
    );

console.log("Graph Start:",minDate);
console.log("Graph End:",maxDate);





    if(tempChart){

        tempChart.destroy();

    }





    tempChart = new Chart(
        document.getElementById("tempChart"),
        {
            type: "line",

            data: {
                datasets: datasets
            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: false,

                interaction: {
                    mode: "nearest",
                    intersect: false
                },

                plugins: {

                    legend: {
                        position: "bottom",
                        labels: {
                            boxWidth: 15
                        }
                    },

                    zoom: {
                        pan: {
                            enabled: true,
                            mode: "x"
                        },

                        zoom: {
                            wheel: {
                                enabled: true
                            },
                            mode: "x"
                        }
                    }

                },


                scales: {

                    x: {
                        type: "time",

                        min: minDate,

                        max: maxDate,

                        ticks: {
                            maxTicksLimit: 7,
                            autoSkip: true,
                            source: "auto"
                        },

                        time: {
                            unit: calculateTimeUnit(minDate, maxDate)
                        }
                    },


                    y: {

                        title: {
                            display: true,
                            text: "Temperature (°C)"
                        },

                        ticks: {
                            maxTicksLimit: 8
                        }

                    }

                }

            }

        }
    );
    document
    .getElementById("chartContainer")
    .classList
    .remove("hidden");



}
// ======================================================
// PART 2 - EXPORT EXCEL REPORT
// ======================================================


const downloadBtn =
document.getElementById("downloadBtn");





if(downloadBtn){

    downloadBtn.addEventListener(
    "click",
    async function(){

        await exportExcelReport();

    });

}







// ======================================================
// Create Final Excel Workbook
// ======================================================


async function exportExcelReport(){


    if(combinedData.length===0){


        alert(
            "Generate graph before downloading."
        );


        return;


    }



    const newWorkbook =
    new ExcelJS.Workbook();




    // --------------------------------------------------
    // Copy Original Worksheets
    // --------------------------------------------------


    workbook.SheetNames.forEach(sheetName=>{


        let oldSheet =
        workbook.Sheets[sheetName];



        let json =
        XLSX.utils.sheet_to_json(
            oldSheet,
            {
                header:1
            }
        );



        let newSheet =
        newWorkbook.addWorksheet(
            sheetName
        );



        json.forEach(row=>{


            newSheet.addRow(row);


        });



    });






    // --------------------------------------------------
    // Combined Temperature Worksheet
    // --------------------------------------------------


    let combinedSheet =
    newWorkbook.addWorksheet(
        "Combined Temperature"
    );




    if(combinedData.length>0){



        let headers =
        Object.keys(
            combinedData[0]
        );



        combinedSheet.addRow(headers);




        combinedData.forEach(row=>{


            combinedSheet.addRow(

                headers.map(header=>

                    row[header]

                )

            );


        });



    }





    // Formatting

    combinedSheet.freezePane = "A2";


    combinedSheet.columns.forEach(column=>{


        column.width=18;


    });






    // --------------------------------------------------
    // Temperature Graph Worksheet
    // --------------------------------------------------


    let graphSheet =
    newWorkbook.addWorksheet(
        "Temperature Graph"
    );



    graphSheet.getCell("A1")
    .value =
    "Temperature Trend Graph";



    graphSheet.getCell("A3")
    .value =
    "Generated from selected inverter temperatures";






    // Convert Chart to Image


    let chartCanvas =
    document.getElementById(
        "tempChart"
    );



    let imageData =
    chartCanvas.toDataURL(
        "image/png"
    );




    let imageId =
    newWorkbook.addImage({

        base64:imageData,

        extension:"png"

    });






    graphSheet.addImage(
        imageId,
        {

            tl:{
                col:0,
                row:4
            },


            ext:{

                width:900,

                height:450

            }


        }

    );







    // --------------------------------------------------
    // Download File
    // --------------------------------------------------


    const buffer =
    await newWorkbook.xlsx.writeBuffer();



    const blob =
    new Blob(
        [buffer],
        {
            type:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    );



    const url =
    URL.createObjectURL(blob);



    const a =
    document.createElement("a");



    a.href=url;


    a.download =
    "Solar_Temperature_Report.xlsx";



    a.click();



    URL.revokeObjectURL(url);



}
// ======================================================
// Convert Excel Timestamp Correctly
// ======================================================

function convertExcelDate(value){


    if(value instanceof Date){

        return value;

    }



    if(typeof value === "number"){


        let parsed =
        XLSX.SSF.parse_date_code(value);



        return new Date(

            parsed.y,

            parsed.m-1,

            parsed.d,

            parsed.H,

            parsed.M,

            parsed.S

        );


    }



    if(typeof value === "string"){


        let date =
        new Date(value);



        if(!isNaN(date.getTime())){

            return date;

        }


    }



    return null;


}
// ======================================================
// Decide X-axis time interval
// ======================================================

function calculateTimeUnit(minDate,maxDate){


    let hours = 
    (
        maxDate - minDate
    )
    /
    (1000 * 60 * 60);



    if(hours <= 24){

        return "hour";

    }



    else if(hours <= 168){

        return "day";

    }



    else{

        return "day";

    }


}

