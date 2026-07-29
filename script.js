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

    sheets.forEach(sheetName=>{

        const sheet = workbook.Sheets[sheetName];

        const json = XLSX.utils.sheet_to_json(sheet,{

            defval:""

        });

        if(json.length===0){

            return;

        }

        const headers = Object.keys(json[0]);

        if(

            headers.includes("Time Stamp") &&

            headers.includes("TEMP1") &&

            headers.includes("TEMP2")

        ){

            workbookData[sheetName]=json;

            createCheckbox(sheetName,"TEMP1");

            createCheckbox(sheetName,"TEMP2");

        }

    });

    if(Object.keys(workbookData).length===0){

        alert("No valid inverter worksheets found.");

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
