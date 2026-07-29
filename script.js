// ======================================================
// Temperature Report Generator
// Part 1 - Read Workbook & Detect Inverters
// ======================================================

let workbook = null;
let workbookData = {};

const HEADER_ROW = 1; // Excel Row 2 (0-based)

const fileInput = document.getElementById("excelFile");
const selectionCard = document.getElementById("selectionCard");
const checkboxContainer = document.getElementById("checkboxContainer");

fileInput.addEventListener("change", loadWorkbook);


// ======================================================
// Read Excel Workbook
// ======================================================

function loadWorkbook(event){

    const file = event.target.files[0];

    if(!file) return;

    workbookData = {};
    checkboxContainer.innerHTML = "";

    const reader = new FileReader();

    reader.onload = function(e){

        const data = new Uint8Array(e.target.result);

        workbook = XLSX.read(data,{
            type:"array"
        });

        detectSheets();

    };

    reader.readAsArrayBuffer(file);

}


// ======================================================
// Detect Valid Worksheets
// ======================================================

function detectSheets(){

    checkboxContainer.innerHTML="";
    workbookData={};


    workbook.SheetNames.forEach(sheetName=>{


        const sheet = workbook.Sheets[sheetName];


        // Read complete sheet as rows
        const rows = XLSX.utils.sheet_to_json(sheet,{
            header:1,
            defval:""
        });


        let headerRow = -1;


        // Search first 10 rows for headers
        for(let i=0; i<Math.min(rows.length,10); i++){


            const headers = rows[i].map(h=>

                h.toString()
                .trim()
                .replace(/\s+/g,"")
                .toUpperCase()

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



        if(headerRow === -1){

            console.log(sheetName,"Header not found");
            return;

        }



        // Convert data using detected header row

        const json = XLSX.utils.sheet_to_json(sheet,{
            range:headerRow,
            defval:""
        });



        workbookData[sheetName]=json;


        createCheckbox(sheetName,"TEMP1");

        createCheckbox(sheetName,"TEMP2");


        console.log(
            sheetName,
            "Header found at Excel row",
            headerRow+1
        );


    });



    if(Object.keys(workbookData).length===0){

        alert("No valid inverter worksheets found.");

        return;

    }


    buildSelectionTable();

    selectionCard.classList.remove("hidden");


}
// ======================================================
// Build Checkbox Table
// ======================================================

function buildSelectionTable(){

    checkboxContainer.innerHTML="";

    const table=document.createElement("table");

    table.style.width="100%";
    table.style.borderCollapse="collapse";

    table.innerHTML=`

        <thead>

            <tr>

                <th style="padding:10px;border-bottom:2px solid #27A5AD;text-align:left;">
                    Inverter
                </th>

                <th style="padding:10px;border-bottom:2px solid #27A5AD;">
                    TEMP1
                </th>

                <th style="padding:10px;border-bottom:2px solid #27A5AD;">
                    TEMP2
                </th>

            </tr>

        </thead>

        <tbody></tbody>

    `;

    const tbody=table.querySelector("tbody");

    Object.keys(workbookData).forEach(inv=>{

        const row=document.createElement("tr");

        row.innerHTML=`

            <td style="padding:12px;border-bottom:1px solid #ddd;">
                ${inv}
            </td>

            <td style="text-align:center;border-bottom:1px solid #ddd;">

                <input
                    type="checkbox"
                    value="${inv}|TEMP1"
                >

            </td>

            <td style="text-align:center;border-bottom:1px solid #ddd;">

                <input
                    type="checkbox"
                    value="${inv}|TEMP2"
                >

            </td>

        `;

        tbody.appendChild(row);

    });

    checkboxContainer.appendChild(table);

}


// ======================================================
// Return Selected Series
// ======================================================

function getSelectedSeries(){

    const selected=[];

    document
        .querySelectorAll("#checkboxContainer input:checked")
        .forEach(box=>{

            const values=box.value.split("|");

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

document.getElementById("generateBtn")
.addEventListener("click",()=>{

    const selected=getSelectedSeries();

    if(selected.length===0){

        alert("Please select at least one temperature.");

        return;

    }

    console.log(selected);

    console.log(workbookData);

    // Part 2 starts here
    // generateWorkbook(selected);

});
