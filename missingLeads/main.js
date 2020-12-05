const main = function () {
  setUpWindowListeners();
  setUpRunButton();
  setWindowVarHolder();
  setUpInputEles();
}

function setUpWindowListeners() {
  window.addEventListener("objsRdy", objectsAreReady);
}

function setUpRunButton() {
  _$("#runButton")[0].addEventListener("click", runButtonClicked);
}

function runButtonClicked() {
  if (window.myVars.runBefore) {
    alert("only once per page sorry");
  }
  turnFilesIntoObj();
}

function objectsAreReady() {
  const cro = window.myVars.caseReportObj;
  const ao = window.myVars.activeObj;
  const nho = window.myVars.noHireObj;
  const co = window.myVars.convertedObj;
  if (!(cro && nho && ao && co)) { return; }
  if (window.myVars.runBefore) { return; }
  window.myVars.runBefore = true;
  window.dispatchEvent(new Event("CreateNewFileForDownload"));
}

function turnFilesIntoObj() {
  if (window.myVars.caseReportFile) {
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(window.myVars.caseReportFile);
    fileReader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, {type : "binary"});
      const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets["Case Report 2.0"]);
      window.myVars.caseReportObj = jsonObj;
      window.dispatchEvent(new Event("objsRdy"));
    }
  }
  if (window.myVars.activeFile) {
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(window.myVars.activeFile);
    fileReader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, {type : "binary"});
      const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"]);
      window.myVars.activeObj = jsonObj;
      window.dispatchEvent(new Event("objsRdy"));
    }
  }
  if (window.myVars.noHireFile) {
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(window.myVars.noHireFile);
    fileReader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, {type : "binary"});
      const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"]);
      window.myVars.noHireObj = jsonObj;
      window.dispatchEvent(new Event("objsRdy"));
    }
  }
  if (window.myVars.convertedFile) {
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(window.myVars.convertedFile);
    fileReader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, {type : "binary"});
      const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"]);
      window.myVars.convertedObj = jsonObj;
      window.dispatchEvent(new Event("objsRdy"));
    }
  }
}

function setWindowVarHolder() {
  if (!window.myVars) {
    window.myVars = {};
  }
}

function setUpInputEles() {
  _$("#caseReportFile")[0].addEventListener("change", (e) => {
    window.myVars["caseReportFile"] = e.target.files[0];
  });
  _$("#activeFile")[0].addEventListener("change", (e) => {
    window.myVars["activeFile"] = e.target.files[0];
  });
  _$("#noHireFile")[0].addEventListener("change", (e) => {
    window.myVars["noHireFile"] = e.target.files[0];
  });
  _$("#convertedFile")[0].addEventListener("change", (e) => {
    window.myVars["convertedFile"] = e.target.files[0];
  });
}


function _$(aString) {
  return document.querySelectorAll(aString);
}


window.onload = main;
