window.addEventListener("CreateNewFileForDownload", function() {
  const nho = window.myVars.noHireObj; //no hire object is nho
  const cro = window.myVars.caseReportObj; //case report object is cro
  // const croM = cro.map(function(ele) {
  //   const pn = ele["CLIENT EMAIL"];
  //   if (typeof(pn) == "string") {
  //     if (pn[0] == "1" ) {
  //       return pn.slice(1).replace(/\D/g,'');
  //     } else {
  //       return pn.replace(/\D/g,'');
  //     }
  //   } else {
  //     return pn;
  //   }
  // });
  // const nhoM = nho.map(function(ele){
  //   const pn = ele["Lead Contact Email"];
  //   if (typeof(pn) =="string") {
  //     if (pn[0] =="1") {
  //       return pn.slice(1).replace(/\D/g,'');
  //     } else {
  //       return pn.replace(/\D/g,'');
  //     }
  //   } else {
  //     return pn;
  //   }
  // });
  const croM = cro.map(function(ele) { //cro mapped
    const email = ele["CLIENT EMAIL"];
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
  const foundArr = [];
  const missingArr = [];
  for (e of nhoM) {
    const croI = croM.indexOf(e);
    if (croI != -1) {
      const nhoI = nhoM.indexOf(e);
      const objToPush = nho[nhoI];
      objToPush["Address"] = cro[croI]["ADDRESS"] + ", " + cro[croI]["CITY, STATE  ZIP"];
      foundArr.push(objToPush);
    } else {
      missingArr.push(e);
    }
  }
  window.myVars.foundArr = foundArr;
  const finSheet = XLSX.utils.json_to_sheet(foundArr);
  const finWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(finWb, finSheet);
  XLSX.writeFile(finWb, "AddressesMaybe.xlsx");
	// const wb = XLSX.utils.json_to_sheet(foundArr, {sheet:"Sheet JS"});
  // XLSX.writeFile(wb, ('SheetJSTableExport.' + ("xlsx" || 'xlsx')));
});
// This could probably be a bit simpler
// A better way to search would be a shrinking
// list of numbers to check against the
// bigger list
