/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const toastList = {}
/* global console, document, Excel, Office */

//Polyfill for toggling element attributes. (Using this polyfill because IExplorer doesn't support "Element.toggleAttribute()")
if (!Element.prototype.toggleAttribute) {
  Element.prototype.toggleAttribute = function(name, force) {
    if(force !== void 0) force = !!force

    if (this.hasAttribute(name)) {
      if (force) return true;

      this.removeAttribute(name);
      return false;
    }
    if (force === false) return false;

    this.setAttribute(name, "");
    return true;
  };
}

function refreshHandler(){

  const refreshFormFields = function (callbackResult) {

    if (callbackResult.status === Office.AsyncResultStatus.Succeeded){

      // We get back the whole settings object.
      const settingsObject = callbackResult.value

      // Then we retrieve our own specific object.
      const formFieldsObject = settingsObject.get("form-fields")
      console.log("[!] Retrieved Settings: ", settingsObject)
      console.log("[!] Retrieved Form Fields: ", formFieldsObject)

      if (formFieldsObject === undefined ) return

      // We update the form fields respectively.
      document.getElementById("username").value = formFieldsObject.username;
      // Uncomment the line below to allow password saving
      // document.getElementById("password").value = formFieldsObject.password;
      document.getElementById("project_id").value = formFieldsObject.project_id
      document.getElementById("strategy").value = formFieldsObject.strategy
    }
  }
  // Call Excel API to give us the stored form fields.
  Office.context.document.settings.refreshAsync(refreshFormFields)
}

function saveHandler(){
  // Retrieve the current formField values
  const formFields = retrieveFormFields()

  // Comment Out the line below to allow password saving
  delete formFields.password; // We don't want to store passwords!!.

  Office.context.document.settings.set("form-fields", formFields)
  Office.context.document.settings.saveAsync()
}

async function checkStatusHandler() {

  const formFields = retrieveFormFields();
  delete formFields.strategy; // Since "strategy" is not required/expected by the API, we remove it

  // We retrieve stored submission key gotten from excel file upload.
  const submissionKey = Office.context.document.settings.get("submission-key");
  console.log("[+] Retrieved Submission Key: " + submissionKey);

  if (submissionKey === null || submissionKey.length < 1){
    showAlert("statusCheckFailed")
    return;
  }

  formFields.file_key = submissionKey.trim();
  // await createWorksheetTab({})

  // Use localhost for testing purposes
  // http://localhost:5000/add/status
  const response = await fetch("https://35.209.133.138:1880/birmind-status", {
    method: 'POST',
    body: formFields,
  });

  const worksheetData = await response.json();

  if (worksheetData.length < 2){
    showAlert("worksheetDataNotFound");
    return;
  }

  showAlert("worksheetDataFound")

  // Call the function to create the worksheet tab "Results".
  await createWorksheetTab(worksheetData);

}


async function createWorksheetTab(worksheetData) {

  // worksheetData = {
  //   "score": 0.9,
  //   "inputs": ["Key A", "Key B", "Key Z"],
  //   "output": "Value A",
  //   "results": [ ["Key A", 0.75], ["Key B", 11760], ["Key Z", 10.5] ],
  //   "entries_score": 0.5,
  //   "optimal_point": 10,
  //   "optimization_output": 0.5
  // }
  // Create the product data rows.
  // const productData = [
  //   ["Almonds", 6, 7.5],
  //   ["Coffee", 20, 34.5],
  //   ["Chocolate", 10, 9.56],
  // ];

  //Code block below is for debugging purposes
  // for (const key in worksheetData) {
  //   if (worksheetData.hasOwnProperty(key)){
  //     console.log(`worksheetData[${key}] = ${worksheetData[key]}`);
  //   }
  // }

  const populateSheet = async function (context) {
    // We add a new worksheet named "Results" here.
    const sheet = context.workbook.worksheets.add("Results");
    // We activate the worksheet
    sheet.activate();
    // sheet.load({})
    await context.sync();
    const rowList = [];

    for (const key in worksheetData) {

      if (worksheetData.hasOwnProperty(key)){
        // console.log(`worksheetData[${key}] = ${worksheetData[key]}`);
        if (typeof(worksheetData[key]) === 'object'){
          if (typeof(worksheetData[key][0]) === 'object'){ // Nested List?
            // const rowBuilder = [];
            rowList.push([key, worksheetData[key][0][0], worksheetData[key][0][1]]); // We insert the first row
            for (let i = 1; i < worksheetData[key].length; i++) {
              rowList.push(['', worksheetData[key][i][0],  worksheetData[key][i][1]]);
            }
          }
          else{
            rowList.push([key, worksheetData[key][0], '']) // We insert the first row
            for (let i = 1; i < worksheetData[key].length; i++) {
              rowList.push(['', worksheetData[key][i],  '']);
            }
          }
        }
        else{
          rowList.push([key, worksheetData[key], '']);
        }

      }

    }

    console.log("[!] RowList: \n", rowList, "\nLength: ", rowList.length);
    const dataRange = sheet.getRange(`A1:C${rowList.length}`); // We compute needed rows to insert our data!.
    dataRange.values = rowList; // We insert data into worksheet!.

  }

  await Excel.run(populateSheet);
}


function retrieveFormFields(){
  const formFields = {}

  const dropdown = document.getElementById("strategy")
  const selectedStrategy = dropdown.value
  formFields.username = document.getElementById("username").value
  formFields.password = document.getElementById("password").value;
  formFields.project_id = document.getElementById("project_id").value;
  formFields.strategy = selectedStrategy !== "strategy" ? selectedStrategy : "";
  // console.log(inputObject)
  return formFields
}

// Controls the Close/Open button functionality.
const formInputFields = ["username", "password", "project_id", "strategy"];
function closeHandler(event) {
  for (let i = 0; i < formInputFields.length; i++) {
    const element = document.getElementById(formInputFields[i]);
    element.toggleAttribute("disabled");
  }
event.target.innerText === "Close" ? event.target.innerText = "Open" : event.target.innerText = "Close";
}

// Show's toast alerts!.
function showAlert(alertToShow) {
  const toast = toastList[alertToShow];
  toast.show();
}

const alertObjects = {
  uploadFailed:"toast-for-upload-failed",
  uploadComplete:"toast-for-upload-success",
  statusCheckFailed:"toast-for-check-status-failed",
  // statusCheckComplete:"toast-for-check-status-success",
  worksheetDataFound: "toast-for-worksheet-data-found",
  worksheetDataNotFound:"toast-for-worksheet-data-not-found"
}

// This block of code is automagically called on Add-In load, it initializes our Add-In, so bootstrapping code should be put here.
 Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    console.log("[!] OnReady! ... ")

    document.getElementById("submit-button").onclick = uploadHandler;
    document.getElementById("save-button").onclick = saveHandler
    document.getElementById("refresh-button").onclick = refreshHandler
    document.getElementById("check-status").onclick = checkStatusHandler;
    document.getElementById("close-button").onclick = closeHandler;

    // We get a reference to each toast button and initialize them all, storing them in a list for later access!.
    for (const alertObjectsKey in alertObjects) {
      if (alertObjects.hasOwnProperty(alertObjectsKey)){
        const toast = document.getElementById(alertObjects[alertObjectsKey])
        toastList[alertObjectsKey] = new bootstrap.Toast(toast);
      }
    }


    // const alertObjects = [
    //   "toast-for-upload-failed",
    //   "toast-for-upload-success",
    //   "toast-for-check-status-failed",
    //   // statusCheckComplete:"toast-for-check-status-success",
    //   "toast-for-worksheet-data-found",
    //   "toast-for-worksheet-data-not-found"
    // ]

    // for (let i = 0; i < alertObjects.length; i++) {
    //   const toast = document.getElementById(alertObjects[i])
    //   // myToast.show();
    //   toastList[alertObjects[i]] = new bootstrap.Toast(toast);
    // }

  }
})


function uploadHandler() {
  console.log("[!] I really do get called!!  [!]")
  // return; { sliceSize:  4194304 }
  // Get all of the content from a PowerPoint, Excel or Word document in N KB chunks of text.
  // But in our own instance, we are getting it all at once (not in chunks)
  Office.context.document.getFileAsync(Office.FileType.Compressed, getFileContents);
}

// Retrieve the current excel file in binary and calls the corresponding upload function.
async function getFileContents(callbackResult) {
  // let fileHandle;
  if (callbackResult.status === Office.AsyncResultStatus.Succeeded){
    const fileHandle = callbackResult.value;

    let byteArrayData;
    // We call the function that performs the actual file retrieval.
    await getFileSlices(fileHandle)
        .then((fileSlices) => {
          if (fileSlices.IsSuccess === true){
            byteArrayData = new Uint16Array(fileSlices.data)
          }
        })
        .catch(error => {
          console.log("[!] Error encountered! from [getFileSlices]!: ", error)
        })

    console.log("[!] Xcel! Data!: ", byteArrayData)

    // Here we call the function that performs file upload, passing it the retrieved excel binary.
    workbookUpload(byteArrayData)

  }
  else{

  }

}

// This function actually does the job of uploading the excel file.
async function workbookUpload(uInt8ArrayData) {

  // We retrieve the stored values.
  // const formFields = retrieveFormFields();

  const formData = new FormData();
  formData.append("parametersObject", JSON.stringify(retrieveFormFields()));
  formData.append("excelFileBinary", uInt8ArrayData);

  // Use localhost for testing purposes.
  // http://localhost:5000/add/submit
  const response = await fetch("https://35.209.133.138:1880/birmind-get", {
    method: 'POST',
    body: formData,
  });

  // We expect the result of the upload to be a submission key.
  const keyString = await response.text();
  console.log("Res: ", keyString);

  if (keyString === null || keyString.length <= 0 ){
    showAlert("uploadFailed");
  }
  else if (keyString.length > 0 ){
    // We save the "submission" key received from the post request above.
    Office.context.document.settings.set("submission-key", keyString.trim());
    Office.context.document.settings.saveAsync();
    console.log("Saved submissionKey: ", keyString);
    showAlert("uploadComplete")
  }

}

// The power engine function that does the actual excel file binary retrieval
function getFileSlices(fileHandle) {

  const sliceCount = fileHandle.sliceCount
  let isError = false;
  
  return new Promise(async (resolve, reject) => {
    let documentFileData = [];
    for (let sliceIndex = 0; (sliceIndex < sliceCount) && !isError; sliceIndex++) {

      const sliceReadPromise = new Promise((sliceResolve, sliceReject) => {

        fileHandle.getSliceAsync(sliceIndex, (asyncResult) => {

          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

            documentFileData = documentFileData.concat(asyncResult.value.data);
            sliceResolve({IsSuccess: true, Data: documentFileData});

          }
          else {
            fileHandle.closeAsync();
            sliceReject({IsSuccess: false, ErrorMessage: `Error in reading the slice: ${sliceIndex} of the document`});
          }

        });

      });

      await sliceReadPromise.catch((error) => {
        isError = true;
        console.log(error);
      });
    }

    if (isError || !documentFileData.length) {
      reject('Error while reading document. Please try it again.');
      return;
    }

    fileHandle.closeAsync();

    console.log("[!] Excel Data!: ", documentFileData)
    resolve({
      IsSuccess: true,
      data: documentFileData
    });
  });

}



/* Base64 string to array encoding */
//
// function uint6ToB64 (nUint6) {
//
//   return nUint6 < 26 ?
//       nUint6 + 65
//       : nUint6 < 52 ?
//           nUint6 + 71
//           : nUint6 < 62 ?
//               nUint6 - 4
//               : nUint6 === 62 ?
//                   43
//                   : nUint6 === 63 ?
//                       47
//                       :
//                       65;
//
// }
//
// function base64EncArr (aBytes) {
//
//   var nMod3 = 2, sB64Enc = "";
//
//   for (var nLen = aBytes.length, nUint24 = 0, nIdx = 0; nIdx < nLen; nIdx++) {
//     nMod3 = nIdx % 3;
//     if (nIdx > 0 && (nIdx * 4 / 3) % 76 === 0) { sB64Enc += "\r\n"; }
//     nUint24 |= aBytes[nIdx] << (16 >>> nMod3 & 24);
//     if (nMod3 === 2 || aBytes.length - nIdx === 1) {
//       sB64Enc += String.fromCharCode(uint6ToB64(nUint24 >>> 18 & 63), uint6ToB64(nUint24 >>> 12 & 63), uint6ToB64(nUint24 >>> 6 & 63), uint6ToB64(nUint24 & 63));
//       nUint24 = 0;
//     }
//   }
//
//   return sB64Enc.substr(0, sB64Enc.length - 2 + nMod3) + (nMod3 === 2 ? '' : nMod3 === 1 ? '=' : '==');
//
// }
//
// function base64Encode(str) {
//   return btoa(encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, function(match, p1) {
//     return String.fromCharCode('0x' + p1);
//   }));
// }
//
// // base64Encode('✓ à la mode'); // "4pyTIMOgIGxhIG1vZGU="
//
// function UTF8ArrToStr (aBytes) {
//
//   let sView = "";
//
//   for (var nPart, nLen = aBytes.length, nIdx = 0; nIdx < nLen; nIdx++) {
//     nPart = aBytes[nIdx];
//     sView += String.fromCharCode(
//         nPart > 251 && nPart < 254 && nIdx + 5 < nLen ? /* six bytes */
//             /* (nPart - 252 << 30) may be not so safe in ECMAScript! So...: */
//             (nPart - 252) * 1073741824 + (aBytes[++nIdx] - 128 << 24) + (aBytes[++nIdx] - 128 << 18) + (aBytes[++nIdx] - 128 << 12) + (aBytes[++nIdx] - 128 << 6) + aBytes[++nIdx] - 128
//             : nPart > 247 && nPart < 252 && nIdx + 4 < nLen ? /* five bytes */
//                 (nPart - 248 << 24) + (aBytes[++nIdx] - 128 << 18) + (aBytes[++nIdx] - 128 << 12) + (aBytes[++nIdx] - 128 << 6) + aBytes[++nIdx] - 128
//                 : nPart > 239 && nPart < 248 && nIdx + 3 < nLen ? /* four bytes */
//                     (nPart - 240 << 18) + (aBytes[++nIdx] - 128 << 12) + (aBytes[++nIdx] - 128 << 6) + aBytes[++nIdx] - 128
//                     : nPart > 223 && nPart < 240 && nIdx + 2 < nLen ? /* three bytes */
//                         (nPart - 224 << 12) + (aBytes[++nIdx] - 128 << 6) + aBytes[++nIdx] - 128
//                         : nPart > 191 && nPart < 224 && nIdx + 1 < nLen ? /* two bytes */
//                             (nPart - 192 << 6) + aBytes[++nIdx] - 128
//                             : /* nPart < 127 ? */ /* one byte */
//                             nPart
//     );
//   }
//
//   return sView;
//
// }


// const base64abc = [
//   "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
//   "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
//   "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m",
//   "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
//   "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/"
// ];
//
// function bytesToBase64(bytes) {
//   let result = '', i, l = bytes.length;
//   for (i = 2; i < l; i += 3) {
//     result += base64abc[bytes[i - 2] >> 2];
//     result += base64abc[((bytes[i - 2] & 0x03) << 4) | (bytes[i - 1] >> 4)];
//     result += base64abc[((bytes[i - 1] & 0x0F) << 2) | (bytes[i] >> 6)];
//     result += base64abc[bytes[i] & 0x3F];
//   }
//   if (i === l + 1) { // 1 octet yet to write
//     result += base64abc[bytes[i - 2] >> 2];
//     result += base64abc[(bytes[i - 2] & 0x03) << 4];
//     result += "==";
//   }
//   if (i === l) { // 2 octets yet to write
//     result += base64abc[bytes[i - 2] >> 2];
//     result += base64abc[((bytes[i - 2] & 0x03) << 4) | (bytes[i - 1] >> 4)];
//     result += base64abc[(bytes[i - 1] & 0x0F) << 2];
//     result += "=";
//   }
//   return result;
// }