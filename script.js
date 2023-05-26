const headRow = document.getElementById("headRow");
var tbody = document.getElementById("tbody");
const cellDiv = document.getElementById("cellNo");
const fontBold = document.getElementById("font-bold")
const fontItalic = document.getElementById("font-italic")
const fontUnderline = document.getElementById("font-underline")
const fontLeftAlign = document.getElementById("font-left-align");
const fontCenterAlign = document.getElementById("font-center-align");
const fontRightAlign = document.getElementById("font-right-align");
const sizeChange = document.getElementById("size");
const fontChange = document.getElementById("font-family");
const copy = document.getElementById("copy")
const cut = document.getElementById("cut")
const paste = document.getElementById("paste")
const textColor = document.getElementById("font-color")
const bgColor = document.getElementById("bg-color")
const textColorLabel = document.querySelector(".text-color-label")
const bgColorLabel = document.querySelector(".bg-color-label")
const fileName= document.getElementById("fileName");
var cellId;
var cutValue = {};
const rows = 100;
const cols = 26;
let matrix = new Array(rows);

function commingsoonFunc(){
    alert("This function of a software is not yet realesed ... will be comming soon")
}

for (let i = 0; i < rows; i++) {
    matrix[i] = new Array(cols);
    for (let j = 0; j < cols; j++) {
        matrix[i][j] = {};
    }
}

function updateMatrix(cellId) {
    let obj = {
        style: cellId.style.cssText,
        text: cellId.innerText,
        id: cellId.id,
    };
    let id = cellId.id.split("");
    let i = id[1] - 1;
    let j = id[0].charCodeAt(0) - 65;
    matrix[i][j] = obj;
}

for (let col = 65; col <= 90; col++) {
    let th = document.createElement("th");
    th.innerText = String.fromCharCode(col);
    headRow.appendChild(th)
}

for (let i = 1; i <= 100; i++) {
    let tr = document.createElement("tr");
    let td = document.createElement("td");
    td.innerText = i;
    tr.appendChild(td);
    for (let col = 65; col <= 90; col++) {
        let editableTd = document.createElement("td")
        editableTd.setAttribute("contenteditable", "true");
        editableTd.setAttribute("id", `${String.fromCharCode(col)}${i}`);
        editableTd.addEventListener("focus", (event) => onFocusFunction(event));
        editableTd.addEventListener("input", (event) => onInputFunction(event));
        tr.appendChild(editableTd)
        tbody.appendChild(tr);
    }
}

let numSheets = 1;
let arrMatrix = [matrix];
let currSheetNum = 1;

function onInputFunction(event) {
    updateMatrix(event.target);
}

function downloadJson() {
    const jsonString = JSON.stringify(matrix);
    const blob = new Blob([jsonString], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "sheet-data.json";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

document.getElementById("jsonFile").addEventListener("change", readJsonFile);

function readJsonFile(event) {
    // here we got our file
    const file = event.target.files[0];
    if (file) {
      // reader object which is instance of FileReader;
      const reader = new FileReader();
  
      reader.onload = function (e) {
        const fileContent = e.target.result;
        try {
          const jsonData = JSON.parse(fileContent);
  
          /// we are not doing anything related to upload;
          matrix = jsonData;
          jsonData.forEach((row) => {
            // cell is cell inside matrix (virtual excel)
            row.forEach((cell) => {
              if (cell.id) {
                // myCell is cell inside my DOM or real Excel
                var myCell = document.getElementById(cell.id);
                myCell.innerText = cell.text;
                myCell.style.cssText = cell.style;
              }
            });
          });

        } catch (err) {
         alert("Error in reading JSON file", err);
        }
      };
      reader.readAsText(file);
    }
}

fontBold.addEventListener("click", () => {
    if (cellId.style.fontWeight == "bold") {
        cellId.style.fontWeight = "normal"
    }
    else {
        cellId.style.fontWeight = "bold"
    }
    updateMatrix(cellId);
})

fontItalic.addEventListener("click", () => {
    if (cellId.style.fontStyle == "italic") {
        cellId.style.fontStyle = "normal"
    }
    else {
        cellId.style.fontStyle = "italic"
    }
    updateMatrix(cellId);
})

fontUnderline.addEventListener("click", () => {
    if (cellId.style.textDecoration == "underline") {
        cellId.style.textDecoration = "none"
    }
    else {
        cellId.style.textDecoration = "underline"
    }
    updateMatrix(cellId);
})

fontLeftAlign.addEventListener("click", () => {
    cellId.style.textAlign = "left"
    updateMatrix(cellId);
})

fontCenterAlign.addEventListener("click", () => {
    cellId.style.textAlign = "center"
    updateMatrix(cellId);
})

fontRightAlign.addEventListener("click", () => {
    cellId.style.textAlign = "right"
    updateMatrix(cellId);
})

sizeChange.addEventListener("change", () => {
    cellId.style.fontSize = sizeChange.value;
    updateMatrix(cellId);
})

fontChange.addEventListener("change", () => {
    cellId.style.fontFamily = fontChange.value;
    updateMatrix(cellId);
})

textColor.addEventListener("change", () => {
    cellId.style.color = textColor.value;
    textColorLabel.style.color = textColor.value;
    updateMatrix(cellId);
})

bgColor.addEventListener("change", () => {
    cellId.style.backgroundColor = bgColor.value;
    bgColorLabel.style.color = bgColor.value;
    updateMatrix(cellId);
})

copy.addEventListener("click", () => {
    cutValue = {
        style: cellId.style.cssText,
        text: cellId.innerText
    };
})

cut.addEventListener("click", () => {
    cutValue = {
        style: cellId.style.cssText,
        text: cellId.innerText
    };
    cellId.style = null;
    cellId.innerText = "";
    updateMatrix(cellId);
})

paste.addEventListener("click", () => {
    if (cutValue.text) {
        cellId.style = cutValue.style;
        cellId.innerText = cutValue.text;
    }
    updateMatrix(cellId);
})

function onFocusFunction(event) {
    cellId = event.target;
    cellDiv.innerText = event.target.id;
} 

document.getElementById("add-sheet-btn").addEventListener("click", () => {
    // logic for adding sheet
  
    /// logic for saving prevSheets
    if (numSheets == 1) {
      var myArr = [matrix];
      localStorage.setItem("ArrMatrix", JSON.stringify(myArr));
    } else {
      var prevSheets = JSON.parse(localStorage.getItem("ArrMatrix"));
      var updatedSheets = [...prevSheets, matrix];
      localStorage.setItem("ArrMatrix", JSON.stringify(updatedSheets));
    }
  
    ///updateMy number of sheets
    numSheets++;
    currSheetNum = numSheets;
  
    // cleanup my virtual memory of excel which is matrix;
    for (let i = 0; i < rows; i++) {
      matrix[i] = new Array(cols);
      for (let j = 0; j < cols; j++) {
        matrix[i][j] = {};
      }
    }
  
    // cleaning up excel sheet in HTML;
    tbody.innerHTML = ``;
  
    for (let row = 1; row <= 100; row++) {
      let tr = document.createElement("tr");
      let th = document.createElement("th");
      th.innerText = row;
      tr.appendChild(th);
      // looping from A to Z;
      for (let col = 1; col <= 26; col++) {
        let td = document.createElement("td");
        td.setAttribute("contenteditable", "true");
        // colRow -> A1, B1, C1, D1,
        td.setAttribute("id", `${String.fromCharCode(col + 64)}${row}`);
        // this is the event listener
        td.addEventListener("focus", (event) => onFocusFunction(event));
        td.addEventListener("input", (event) => onInputFunction(event));
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
  
    document.getElementById("sheet-num").innerText = "Sheet No. " + currSheetNum;
  });
  
  document.getElementById("sheet-1").addEventListener("click", () => {
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[0];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });
  
  document.getElementById("sheet-2").addEventListener("click", () => {
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[1];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });
  
  document.getElementById("sheet-3").addEventListener("click", () => {
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[2];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });