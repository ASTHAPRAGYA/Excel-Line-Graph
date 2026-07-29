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

                new Date(
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

            new Date(
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



    selected.forEach((item,index)=>{


        let key =
        item.inverter +
        " " +
        item.temperature;



        let dataPoints = combinedData
        .map(row=>{


            return {

                x: row["Time Stamp"],

                y: row[key]

            };


        })
        .filter(point=>

            point.y !== "" &&
            point.y !== null &&
            point.y !== undefined &&
            !isNaN(point.y)

        );




        datasets.push({


            label:key,


            data:dataPoints,


            borderColor:
            colors[index % colors.length],


            backgroundColor:
            colors[index % colors.length],



            borderWidth:2,


            pointRadius:0,


            pointHoverRadius:5,


            tension:0.15,


            fill:false



        });



    });





    if(tempChart){

        tempChart.destroy();

    }






    tempChart = new Chart(

        document
        .getElementById("tempChart"),


        {


        type:"line",



        data:{


            datasets:datasets


        },



        options:{


            responsive:true,


            maintainAspectRatio:false,



            parsing:false,



            animation:false,



            interaction:{


                mode:"nearest",

                intersect:false


            },



            plugins:{


                legend:{


                    position:"top"



                },



                tooltip:{


                    callbacks:{


                        title:function(context){


                            return new Date(
                                context[0].parsed.x
                            )
                            .toLocaleString();


                        },


                        label:function(context){


                            return (

                                context.dataset.label
                                +
                                ": "
                                +
                                context.parsed.y
                                +
                                " °C"

                            );


                        }


                    }


                },



                zoom:{


                    pan:{


                        enabled:true,

                        mode:"x"


                    },



                    zoom:{


                        wheel:{


                            enabled:true


                        },


                        pinch:{


                            enabled:true


                        },


                        mode:"x"


                    }


                }


            },



            scales:{


                x:{


                    type:"time",



                    time:{


                        tooltipFormat:
                        "dd MMM yyyy HH:mm"



                    },



                    ticks:{


                        maxTicksLimit:15


                    }



                },




                y:{


                    title:{


                        display:true,


                        text:
                        "Temperature (°C)"


                    },


                    ticks:{


                        callback:function(value){


                            return value+" °C";


                        }


                    }



                }



            }



        }



    });






    document
    .getElementById("chartContainer")
    .classList
    .remove("hidden");



}
