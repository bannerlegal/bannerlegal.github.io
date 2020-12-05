window.addEventListener("CreateNewFileForDownload", function() {
  const ao = window.myVars.activeObj; //active object is ao
  const nho = window.myVars.noHireObj; //no hire object is nho
  const co = window.myVars.convertedObj; //converted object is co
  const cro = window.myVars.caseReportObj; //case report object is cro
  const croM = cro.map(function(ele) { //cro mapped
    const email = ele["CLIENT EMAIL"];
    const isVa = ele["VA"];
    if (isVa == undefined) {
      return "no";
    }
    if (typeof(email) == "string") {
      return email.toLowerCase().trim();
    } else {
      return email;
    }
  });
  const aoM = ao.map(function(ele) { //nho mapped
    const email = ele["Lead Contact Email"];
    if (typeof(email) == "string") {
      return email.toLowerCase().trim();
    } else {
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
  const coM = co.map(function(ele) { //nho mapped
    const email = ele["Lead Contact Email"];
    if (typeof(email) == "string") {
      return email.toLowerCase().trim();
    } else {
      return email;
    }
  });
  const foundArr = [];
  const missingArr = [];
  for (e of croM) {
    if (e == "no") {
      continue;
    }
    const aoI = aoM.indexOf(e);
    const nhoI = nhoM.indexOf(e);
    const coI = coM.indexOf(e);
    if (aoI != -1) {

    } else if (nhoI != -1) {

    } else if (coI != -1) {

    } else {
      const croI = croM.indexOf(e);
      missingArr.push(cro[croI]);
    }
  }
  // for (e of nhoM) {
  //   const croI = croM.indexOf(e);
  //   if (croI != -1) {
  //     const nhoI = nhoM.indexOf(e);
  //     const objToPush = nho[nhoI];
  //     objToPush["ADDRESS"] = cro[croI]["ADDRESS"];
  //     objToPush["CITY, STATE  ZIP"] = cro[croI]["CITY, STATE  ZIP"];
  //     foundArr.push(objToPush);
  //   } else {
  //     missingArr.push(e);
  //   }
  // }
  window.myVars.foundArr = foundArr;
  window.myVars.missingArr = missingArr;
  const finSheet = XLSX.utils.json_to_sheet(missingArr);
  const finWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(finWb, finSheet);
  XLSX.writeFile(finWb, "missingEntries.xlsx");
	// const wb = XLSX.utils.json_to_sheet(foundArr, {sheet:"Sheet JS"});
  // XLSX.writeFile(wb, ('SheetJSTableExport.' + ("xlsx" || 'xlsx')));
});
// This could probably be a bit simpler
// A better way to search would be a shrinking
// list of numbers to check against the
// bigger list
