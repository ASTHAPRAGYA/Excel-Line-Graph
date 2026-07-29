// ======================================================
// Solar Inverter Temperature Analyzer
// Complete Script
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
// Upload Excel
// ======================================================


fileInput.addEventListener("change", function(event){


    const file = event.target.files[0];


    if(!file)
        return;



    const reader = new FileReader();



    reader.onload = function(e){


        const data = new Uint8Array(e.target.result);



        workbook = XLSX.read(data,{

            type:"array"

        });



        console.log(
            "Workbook Loaded",
            workbook.SheetNames
        );



        detectSheets();



    };



    reader.readAsArrayBuffer(file);



});







// ======================================================
// Detect Sheets + Header Rows
// ======================================================


function detectSheets(){


    workbookData = {};

    checkboxContainer.innerHTML="";



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

            "Header row:",

            headerRow+1

        );





        let rawData =
        XLSX.utils.sheet_to_json(sheet,{

            range:headerRow,

            defval:""

        });





        let cleanedData =
        cleanSheetData(rawData);





        if(cleanedData.length>0){


            workbookData[sheetName]=cleanedData;


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
            "No inverter worksheets found."
        );

        return;

    }




    buildSelectionTable();



    selectionCard.classList.remove(
        "hidden"
    );



}








// ======================================================
// Normalize Header Names
// ======================================================


function normalizeHeader(value){


    return value
    .toString()
    .trim()
    .replace(/\s+/g,"")
    .toUpperCase();


}







// ======================================================
// Clean Data
// Converts headers into standard format
// ======================================================


function cleanSheetData(data){



    return data.map(row=>{


        let newRow={};



        Object.keys(row).forEach(key=>{


            let clean =
            normalizeHeader(key);



            if(clean==="TIMESTAMP"){

                newRow["Time Stamp"]
                =
                row[key];

            }



            else if(clean==="TEMP1"){

                newRow["TEMP1"]
                =
                Number(row[key]);

            }



            else if(clean==="TEMP2"){

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
// Build Selection Table
// ======================================================


function buildSelectionTable(){


    checkboxContainer.innerHTML="";



    let table=document.createElement("table");



    table.style.width="100%";

    table.innerHTML=`

    <tr>

    <th>
    Inverter
    </th>

    <th>
    TEMP1
    </th>

    <th>
    TEMP2
    </th>


    </tr>

    `;





    Object.keys(workbookData)
    .forEach(inv=>{


        let row=document.createElement("tr");



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
// Get Selected Channels
// ======================================================


function getSelectedSeries(){


    let selected=[];



    document
    .querySelectorAll(
        "#checkboxContainer input:checked"
    )
    .forEach(box=>{


        let values =
        box.value.split("|");



        selected.push({

            inverter:values[0],

            temperature:values[1]

        });



    });



    return selected;


}









// ======================================================
// Generate Button
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



}

);









// ======================================================
// Create Combined Dataset
// ======================================================


function createCombinedData(){



    combinedData=[];



    let timeSet=new Set();





    Object.keys(workbookData)
    .forEach(inv=>{


        workbookData[inv]
        .forEach(row=>{


            timeSet.add(

                new Date(
                    row["Time Stamp"]
                ).getTime()

            );


        });


    });






    let times =
    Array.from(timeSet)
    .sort(
        (a,b)=>a-b
    );





    times.forEach(time=>{


        let obj={};


        obj["Time Stamp"]=
        new Date(time);




        Object.keys(workbookData)
        .forEach(inv=>{


            obj[inv+" TEMP1"]="";

            obj[inv+" TEMP2"]="";


        });



        combinedData.push(obj);



    });






    Object.keys(workbookData)
    .forEach(inv=>{


        workbookData[inv]
        .forEach(row=>{


            let t =
            new Date(
                row["Time Stamp"]
            ).getTime();



            let index =
            times.indexOf(t);




            if(index>=0){


                combinedData[index]
                [inv+" TEMP1"]
                =
                row["TEMP1"];



                combinedData[index]
                [inv+" TEMP2"]
                =
                row["TEMP2"];



            }



        });



    });



}









// ======================================================
// Create Chart
// ======================================================


function createChart(selected){



    let datasets=[];



    selected.forEach(item=>{


        let key =
        item.inverter+
        " "+
        item.temperature;




        datasets.push({

            label:key,


            data:

            combinedData.map(row=>({

                x:row["Time Stamp"],

                y:row[key]

            })),



            borderWidth:2,

            pointRadius:0,

            tension:0.1



        });



    });





    if(tempChart){

        tempChart.destroy();

    }




    let ctx =
    document
    .getElementById("tempChart")
    .getContext("2d");






    tempChart =
    new Chart(ctx,{


        type:"line",



        data:{


            datasets:datasets


        },




        options:{


            responsive:true,


            maintainAspectRatio:false,



            parsing:false,



            interaction:{


                mode:"nearest",

                intersect:false


            },



            plugins:{


                legend:{

                    position:"top"

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


                        unit:
                        calculateTimeUnit()


                    }



                },




                y:{


                    title:{


                        display:true,

                        text:
                        "Temperature (°C)"

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









function calculateTimeUnit(){


    if(combinedData.length===0)

        return "hour";



    let days =

    (
        combinedData[
        combinedData.length-1
        ]["Time Stamp"]

        -

        combinedData[0]
        ["Time Stamp"]

    )

    /
    (1000*60*60*24);





    if(days<=1)

        return "hour";


    if(days<=7)

        return "6hour";



    return "day";


}
