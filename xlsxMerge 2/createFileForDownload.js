window.addEventListener("CreateNewFileForDownload", function() {
  console.warn("Clear to launch");
  const lo = window.myVars.leadsObj;
  const cro = window.myVars.caseReportObj; //case report object is cro
  const caseNames = lo.map(e => e["Case Name"]);
  let arrOfNamesToFind = caseNames.map(e => e.match(/(\w{3,30})[^(]+?(\w{3,30})/i));
  const croNameAndIndex = cro.map((e, i) => [e["FIRST NAME"]+e["LAST NAME"], i]);
  arrOfNamesToFind = arrOfNamesToFind.filter(e => e);
  arrOfNamesToFind = arrOfNamesToFind.filter(e => e[1] && e[0]);
  const indexOfNames = arrOfNamesToFind.map(ne => {
    const matchFirstName = croNameAndIndex.filter(ce => (new RegExp(ne[1], "i")).test(ce[0]));
    const matchLastNameToo = croNameAndIndex.filter(ce => (new RegExp(ne[2], "i")).test(ce[0]));
    if (matchLastNameToo.length == 0) {
      return {"FIRST NAME": `COULD NOT FIND: ${ne[0]}`}
    }
    return matchLastNameToo[0][1];
  });
  const processedArr = indexOfNames.map(e => {
    if (typeof e == "object") {
      return e;
    }
    return cro[e];
  });

  const foundArr = processedArr.filter(e => !/find/i.test(e["FIRST NAME"]));
  const missingArr = processedArr.filter(e => /find/i.test(e["FIRST NAME"]));

  missingArr.forEach(e => {
    const eContainer = document.createElement("div");
    eContainer.innerText = e["FIRST NAME"];
    eContainer.classList.add("missingPerson");
    document.querySelector("body").appendChild(eContainer);
  });
  console.warn("We have lift off");
  const finSheet = XLSX.utils.json_to_sheet(foundArr);
  const finWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(finWb, finSheet);
  XLSX.writeFile(finWb, (new Date()).toLocaleDateString("en-US").replaceAll("/", "-")+"-foundArr.xlsx");
});
