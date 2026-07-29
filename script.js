// ===============================
// Temperature Report Generator
// Part 1
// ===============================

let workbook = null;
let workbookData = {};

const fileInput = document.getElementById("excelFile");
const checkboxContainer = document.getElementById("checkboxContainer");
const selectionCard = document.getElementById("selectionCard");

fileInput.addEventListener("change", loadWorkbook);


//----------------------------------------
// Read Uploaded Workbook
//----------------------------------------

async function loadWorkbook(event){

    const file = event.target.files[0];

    if(!file){

        return;

    }

    checkboxContainer.innerHTML="";

    workbookData={};

    const reader = new FileReader();

    reader.onload = async function(e){

        const data = new Uint8Array(e.target.result);

        workbook = XLSX.read(data,{

            type:"array"

        });

        detectSheets();

    };

    reader.readAsArrayBuffer(file);

}


//----------------------------------------
// Detect worksheets automatically
//----------------------------------------

function detectSheets(){

    const sheets = workbook.SheetNames;

    checkboxContainer.innerHTML="";
    workbookData={};

    sheets.forEach(sheetName=>{

        const ws = workbook.Sheets[sheetName];

        // Convert sheet to array instead of JSON
        const rows = XLSX.utils.sheet_to_json(ws,{
            header:1,
            defval:""
        });

        if(rows.length===0) return;

        let headerRow=-1;

        // Search first 20 rows for headers
        for(let r=0;r<Math.min(20,rows.length);r++){

            const headers=rows[r].map(x=>
                x.toString()
                 .trim()
                 .replace(/\s+/g,"")
                 .toUpperCase()
            );

            if(
                headers.includes("TIMESTAMP") &&
                headers.includes("TEMP1") &&
                headers.includes("TEMP2")
            ){

                headerRow=r;
                break;

            }

        }

        if(headerRow==-1){

            console.log(sheetName+" skipped");
            return;

        }

        const json=XLSX.utils.sheet_to_json(ws,{
            range:headerRow,
            defval:""
        });

        workbookData[sheetName]=json;

        createCheckbox(sheetName,"TEMP1");
        createCheckbox(sheetName,"TEMP2");

    });

    console.log(workbookData);

    if(Object.keys(workbookData).length===0){

        alert("No valid worksheets found.");

        return;

    }

    selectionCard.classList.remove("hidden");

}

//----------------------------------------
// Create Checkbox
//----------------------------------------

function createCheckbox(inv,temp){

    const div=document.createElement("div");

    div.className="checkboxItem";

    const id=inv+"_"+temp;

    div.innerHTML=`

        <label>

            <input

                type="checkbox"

                value="${id}"

            >

            ${inv} ${temp}

        </label>

    `;

    checkboxContainer.appendChild(div);

}


//----------------------------------------
// Utility
//----------------------------------------

function getSelectedSeries(){

    const checked=[];

    document

        .querySelectorAll("#checkboxContainer input:checked")

        .forEach(box=>{

            checked.push(box.value);

        });

    return checked;

}
