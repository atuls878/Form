window.addEventListener("load", function () {
  fetch("./po.xlsx")
    .then((response) => response.blob())
    .then((blob) => readXlsxFile(blob))
    .then((rows) => {
      updateDropDown(rows);
    });
});

let obj = {};
let poObj = {};

function updateDropDown(rows) {
  let suppliers = [];

  for (let i = 1; i < rows.length; i++) {
    let supplierName = rows[i][11],
      poNumber = rows[i][3],
      description = rows[i][15];
    if (
      supplierName != "" &&
      supplierName != null &&
      supplierName != undefined &&
      suppliers.indexOf(supplierName) == -1
    ) {
      suppliers.push(supplierName);
    }

    if (poNumber != "" && poNumber != null && poNumber != undefined) {
      if (obj[supplierName] == undefined) {
        obj[supplierName] = [];
        obj[supplierName].push(poNumber);
      }
      if (obj[supplierName].indexOf(poNumber) == -1) {
        obj[supplierName].push(poNumber);
      }
      poObj[poNumber] = description;
    }
  }
  suppliers.map((value) => updateOptions("supplier", value));
}

let supplier = document.querySelector("#supplier");
let poNumber = document.querySelector("#poNumber");

supplier.addEventListener("change", (e) => {
  let selectedSupplierName = e.target.value;
  let poArr = obj[selectedSupplierName];
  poArr.unshift("Select PO Number");
  poNumber.innerHTML = null;
  poArr.map((value) => updateOptions("poNumber", value));
});

function updateOptions(query, option) {
  let supplier = document.querySelector(`#${query}`);
  let ele = document.createElement("option");
  ele.innerHTML = option;
  supplier.appendChild(ele);
}

let form = document.querySelector("#formComponent");

let fullName = document.querySelector("#name");
let startTime = document.querySelector("#startTime");
let endTime = document.querySelector("#endTime");
let hours = document.querySelector("#hours");
let rate = document.querySelector("#rate");

form.addEventListener("submit", (e) => {
  e.preventDefault();

  const writeObj = [
    {
      name: "",
      startTime: "",
      endTime: "",
      hours: "",
      rate: "",
      supplier: "",
      po: "",
      description: "",
    },
  ];

  let nameValue = fullName.value;
  let startValue = startTime.value;
  let endValue = endTime.value;
  let hourValue = hours.value;
  let rateValue = rate.value;
  let supplierValue = supplier.value;
  let poValue = poNumber.value;

  writeObj[0].name = nameValue;
  writeObj[0].startTime = startValue;
  writeObj[0].endTime = endValue;
  writeObj[0].hours = hourValue;
  writeObj[0].rate = rateValue;
  writeObj[0].supplier = supplierValue;
  writeObj[0].po = poValue;
  writeObj[0].description = poObj[poValue];

  const schema = [
    {
      column: "Name",
      type: String,
      value: (writeObj) => writeObj.name,
    },
    {
      column: "Start Time",
      type: String,
      value: (writeObj) => writeObj.startTime,
    },
    {
      column: "End Time",
      type: String,
      value: (writeObj) => writeObj.endTime,
    },
    {
      column: "Hours Worked",
      type: String,
      value: (writeObj) => writeObj.hours,
    },
    {
      column: "Rate per Hour",
      type: String,
      value: (writeObj) => writeObj.rate,
    },
    {
      column: "Supplier Name",
      type: String,
      value: (writeObj) => writeObj.supplier,
    },
    {
      column: "Purchase Order",
      type: String,
      value: (writeObj) => writeObj.po,
    },
    {
      column: "Description",
      type: String,
      value: (writeObj) => writeObj.description,
    },
  ];

  // console.log(writeObj);

  writeXlsxFile(writeObj, {
    schema,
    fileName: "output.xlsx",
  });
});
