window.addEventListener("CreateNewFileForDownload", function() {
  console.warn("Clear to launch");
  const nho = window.myVars.noHireObj; //no hire object is nho
  const cro = window.myVars.caseReportObj; //case report object is cro
  const croM = cro.map(function(ele) { //cro mapped
    const email = ele["CLIENT EMAIL"];
    if (typeof(email) == "string") {
      return email.toLowerCase().trim();
    } else {
      console.log(`These emails are not strings:\n ${email}\n`);
      return email;
    }
  });
  const nhoM = nho.map(function(ele) { //nho mapped
    const email = ele["Lead Contact Email"];
    if (typeof(email) == "string") {
      return email.toLowerCase().trim();
    } else {
      return email;
    }
  });
  const foundArr = [];
  const missingArr = [];
  for (e of nhoM) {
    const croI = croM.indexOf(e);
    if (croI != -1) {
      const nhoI = nhoM.indexOf(e);


      // const objToPush = nho[nhoI];
      // objToPush["ADDRESS"] = cro[croI]["ADDRESS"];
      // objToPush["CITY, STATE  ZIP"] = cro[croI]["CITY, STATE  ZIP"];
      // delete objToPush["Lead Name"];
      // objToPush["First Name"] = cro[croI]["FIRST NAME"];
      // objToPush["Last Name"] = cro[croI]["LAST NAME"];


      const objToPush = {};
      objToPush["First Name"] = cro[croI]["FIRST NAME"];
      objToPush["Last Name"] = cro[croI]["LAST NAME"];
      objToPush["ADDRESS"] = cro[croI]["ADDRESS"];
      objToPush["CITY, STATE  ZIP"] = cro[croI]["CITY, STATE  ZIP"];
      objToPush["No Hire Reason"] = nho[nhoI]["No Hire Reason"];
      objToPush["No Hire Date"] = nho[nhoI]["No Hire At"];


      foundArr.push(objToPush);
    } else {
      missingArr.push(e);
    }
  }
  window.myVars.foundArr = foundArr;
  console.warn("We have lift off");
  const finSheet = XLSX.utils.json_to_sheet(foundArr);
  const finWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(finWb, finSheet);
  XLSX.writeFile(finWb, (new Date()).toLocaleDateString("en-US").replaceAll("/", "-")+"-NoHireAddress.xlsx");
	// const wb = XLSX.utils.json_to_sheet(foundArr, {sheet:"Sheet JS"});
  // XLSX.writeFile(wb, ('SheetJSTableExport.' + ("xlsx" || 'xlsx')));
});
// This could probably be a bit simpler
// A better way to search would be a shrinking
// list of numbers to check against the
// bigger list
