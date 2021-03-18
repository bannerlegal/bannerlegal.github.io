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
    console.log("only once per page sry");
    return;
  }
  turnFilesIntoObj();
}

function objectsAreReady() {
  const cro = window.myVars.caseReportObj;
  const lo = window.myVars.leadsObj;
  if (!(cro && lo)) return;
  if (window.myVars.runBefore) return;
  console.log("Hit");
  window.dispatchEvent(new Event("CreateNewFileForDownload"));
  window.myVars.runBefore = true;
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
  if (window.myVars.leadsFile) {
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(window.myVars.leadsFile);
    fileReader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, {type : "binary"});
      const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
      window.myVars.leadsObj = jsonObj;
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
  _$("#leadsFile")[0].addEventListener("change", (e) => {
    window.myVars["leadsFile"] = e.target.files[0];
  });
}


function _$(aString) {
  return document.querySelectorAll(aString);
}


window.onload = main;
