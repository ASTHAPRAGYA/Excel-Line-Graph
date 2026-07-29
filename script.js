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

    workbook.SheetNames.forEach(sheetName=>{

        const sheet = workbook.Sheets[sheetName];

        const json = XLSX.utils.sheet_to_json(sheet,{
            range:HEADER_ROW,
            defval:""
        });

        if(json.length===0)
            return;

        const headers = Object.keys(json[0]).map(h=>
            h.toString().trim().toUpperCase()
        );

        const hasTime = headers.includes("TIME STAMP");
        const hasTemp1 = headers.includes("TEMP1");
        const hasTemp2 = headers.includes("TEMP2");

        if(hasTime && hasTemp1 && hasTemp2){

            workbookData[sheetName]=json;

        }

    });

    if(Object.keys(workbookData).length===0){

        alert("No inverter worksheets found.");

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
