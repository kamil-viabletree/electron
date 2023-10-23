const XLSX = require("xlsx");
const Swal = require("sweetalert2");
const path = require("path");
const fs = require("fs");

document.addEventListener("DOMContentLoaded", function () {
  sequenceNameGenerator();
  var add_field = document.getElementById("add-field");

  // Add a click event listener dynamic-fields
  add_field.addEventListener("click", function () {
    dynamicFieldAdd();
  });

  function dynamicFieldAdd() {
    const uuid = new Date().getTime();
    const count = document.querySelectorAll(".sequence-field").length;
    const name = sequenceNameGenerator(count+1);
    const fieldHTML = `<div class="row sequence-field" id="${uuid}_field">
        <div class="col-10">
            <div class="form-group">
                <label for="${uuid}">${name} To Last</label>
                <input
                    pattern="^(?!\s*$).+"
                    required
                    type="text"
                    class="form-control sequence-field-value"
                    placeholder="Enter ${name} To Last"
                    id="${uuid}"
                />
            </div>
        </div>
        <div class="col-2">
            <button type="button" class="btn btn-danger font-weight-bolder del-btn" id="${uuid}"><i class="fa fa-trash-alt"></i></button>
        </div>
    </div>`;

    // Create a temporary container element
    const tempContainer = document.createElement("div");
    tempContainer.className = `row `;
    tempContainer.innerHTML = fieldHTML;

    // Get the actual field element from the container
    const fieldElement = tempContainer.firstElementChild;

    // Append the new field to the dynamic-fields container
    document.querySelector(".dynamic-fields").appendChild(fieldElement);
    deleteFieldsWithUUID();
  }

  function deleteFieldsWithUUID() {
    var clickableDivs = document.querySelectorAll(".del-btn");
    if (clickableDivs != null) {
      clickableDivs.forEach(function (div) {
        div.addEventListener("click", function () {
          var clickedID = div.id;
          let field = document.getElementById(`${clickedID}_field`);
          if (field) {
            field.remove();
          }
        });
      });
    }
  }

  function createDirectoryIfNotExists(directoryPath) {
    if (!fs.existsSync(directoryPath)) {
      fs.mkdirSync(directoryPath, { recursive: true });
    }
  }

  function sequenceNameGenerator(n) {
    var special = [
      "Zeroth",
      "First",
      "Second",
      "Third",
      "Fourth",
      "Fifth",
      "Sixth",
      "Seventh",
      "Eighth",
      "Ninth",
      "Tenth",
      "Eleventh",
      "Twelfth",
      "Thirteenth",
      "Fourteenth",
      "Fifteenth",
      "Sixteenth",
      "Seventeenth",
      "Eighteenth",
      "Nineteenth",
    ];
    var deca = [
      "Twent",
      "Thirt",
      "Fort",
      "Fift",
      "Sixt",
      "Sevent",
      "Eight",
      "Ninet",
    ];

    if (n < 20) return special[n];
    if (n % 10 === 0) return deca[Math.floor(n / 10) - 2] + "ieth";
    return deca[Math.floor(n / 10) - 2] + "y " + special[n % 10];
  }

  const downloadDir = __dirname.match(/^[A-Z]:/i)[0];
  const directoryPath = path.join(downloadDir, "Konica");
  let dir =
    "E:/Stupell Industries Dropbox/Manufacturing and Printing/Shared Picklists/ReRunPrep Picklists";
  createDirectoryIfNotExists(directoryPath);

  const storagePath = localStorage.getItem("elec-storage-path");

  if (storagePath === null) {
    localStorage.setItem("elec-storage-path", dir);
  } else {
    dir = storagePath;
  }
  document.getElementById("storage-path").value = dir.replace(/\//g, "\\");
  document.getElementById("storage-path-button").disabled = true;

  document
    .getElementById("storage-path")
    .addEventListener("keyup", function (e) {
      const newPath = document.getElementById("storage-path").value;
      if (newPath == dir.replace(/\//g, "\\")) {
        document.getElementById("storage-path-button").disabled = true;
      } else {
        document.getElementById("storage-path-button").disabled = false;
      }
    });

  document
    .getElementById("storage-path-button")
    .addEventListener("click", function (e) {
      const newPath = document.getElementById("storage-path").value;
      if (newPath.trim() == "") {
        Swal.fire({
          title: "Info!",
          text: "Path field should be pre filled!",
          icon: "info",
          confirmButtonText: "OK",
        }).then((e) => {
          if (e.isConfirmed) {
            document.getElementById("storage-path").value = dir.replace(
              /\//g,
              "\\"
            );
            document.getElementById("storage-path-button").disabled = true;
          }
        });
      } else {
        localStorage.setItem(
          "elec-storage-path",
          newPath.replace(/\\/g, "/").trim()
        );
        dir = newPath.replace(/\\/g, "/").trim();
        document.getElementById("storage-path-button").disabled = true;
      }
    });

  document.getElementById("myForm").addEventListener("submit", function (e) {
    e.preventDefault();
    const sequencrArr = [];
    const field_values = document.querySelectorAll(".sequence-field-value");
    const field_count = field_values.length;
    field_values.forEach(e=>{
      if(e.value.trim() != ""){
        sequencrArr.push(e.value.trim())
      }
    });
    const excel_file = document.getElementById("excel-file");

    if (excel_file.files.length > 0) {
      const selectedFile = excel_file.files[0];
      const selectedFileName = selectedFile.name;
      const pattern = /statusReport_/;
      const match = pattern.exec(selectedFileName);

      if (match) {
        // const fileName = match[1];
        const fileName = selectedFileName.replace("statusReport", "RR");
        const reader = new FileReader();

        reader.onload = function (e) {
          const data = e.target.result;

          // const workbook = XLSX.read(data, { type: "binary" });
          processFile(data, sequencrArr, field_count, fileName, dir);
        };
        reader.readAsBinaryString(selectedFile);
      } else {
        Swal.fire({
          title: "Error!",
          text: "Wrong file",
          icon: "error",
          confirmButtonText: "OK",
        });
      }
    } else {
      Swal.fire({
        title: "Error!",
        text: "Konica Print sheet not found!",
        icon: "error",
        confirmButtonText: "OK",
      });
    }

  });

  function processFile(data, sequencrArr, field_count, fileName, dir) {
    let storeFile = fileName;

    const workbook = XLSX.read(data, { type: "binary" });

    const sheetNameIndex = workbook.SheetNames.indexOf("KonicaPrint");

    // KonicaPrint
    if (sheetNameIndex !== -1) {
      const sheet = workbook.Sheets[workbook.SheetNames[sheetNameIndex]];
      if (sheet) {
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        let count = field_count;
        const printFiles = [];
        const missingFiles = [];

        for (let i = 1; i < jsonData.length; i++) {
          if (jsonData[i]["Missing Files Konica"]) {
            missingFiles.push({
              ProductID: jsonData[i]["Missing Files Konica"],
              Qty: jsonData[i]["__EMPTY_2"],
              RunningTally: "",
              OrderID: "",
              KonicaMimaki: "",
            });
          }
          if (jsonData[i]["Ready to Print Konica"]) {
            if (count == 0) {
              printFiles.push({
                ProductID: jsonData[i]["Ready to Print Konica"],
                Qty: jsonData[i]["__EMPTY"],
                RunningTally: "",
                OrderID: "",
                KonicaMimaki: "",
              });
            } else {
              if (sequencrArr.includes(jsonData[i]["Ready to Print Konica"])) {
                count--;
              } else {
                if (field_count != 0 && count <= field_count - 1) {
                  count++;
                }
              }
            }
          }
        }

        if (printFiles.length > 0 || missingFiles.length > 0) {
          createNewSheet(printFiles, missingFiles, storeFile, dir);
        } else {
          // error
          Swal.fire({
            title: "Info!",
            text: "No remmainings were found",
            icon: "info",
            confirmButtonText: "OK",
          });
        }
      } else {
        Swal.fire({
          title: "Info!",
          text: "No data found",
          icon: "info",
          confirmButtonText: "OK",
        });
      }
    } else {
      Swal.fire({
        title: "Error!",
        text: "KonicaPrint sheet not found",
        icon: "error",
        confirmButtonText: "OK",
      });
    }
  }

  function isFileOpen(filePath) {
    try {
      // Try to open the file in write mode, which will throw an error if it's already open.
      fs.openSync(filePath, "w");
      return false; // The file is not open by another application.
    } catch (error) {
      return true; // The file is open by another application.
    }
  }

  /**
   * New Sheet creating
   */

  function createNewSheet(printFiles, missingFiles, storeFile, dir) {
    // Create a new newWorkbook
    const newWorkbook = XLSX.utils.book_new();

    // Create a new worksheet
    const worksheet = XLSX.utils.json_to_sheet([
      ...printFiles,
      ...missingFiles,
    ]);

    // Add the worksheet to the newWorkbook
    XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Sheet1");

    // Write the newWorkbook to a file
    let downloadDir = "";
    let directoryPath = "";
    if (fs.existsSync(dir)) {
      // If directory exists
      directoryPath = path.join(dir, `${storeFile}`);
    } else {
      downloadDir = __dirname.match(/^[A-Z]:/i)[0];
      directoryPath = path.join(downloadDir, `Konica/${storeFile}`);
    }

    if (isFileOpen(directoryPath)) {
      Swal.fire({
        title: "Write Error!",
        text: `This is open ${directoryPath}, please close to proceed`,
        icon: "error",
        confirmButtonText: "Ok",
      });
    } else {
      XLSX.writeFile(newWorkbook, directoryPath);
      document.getElementById("myForm").reset();

      if (fs.existsSync(dir)) {
        // Trigger a click event on the anchor element to initiate the download
        Swal.fire({
          title: "Processed Successfully!",
          text: `Your file is stored in this ${directoryPath}`,
          icon: "success",
          confirmButtonText: "Ok",
        });
      } else {
        Swal.fire({
          title: "Processed Successfully!",
          text: `Your file is ready to download`,
          showCancelButton: true,
          confirmButtonText: "Download",
          denyButtonText: `Cancel`,
        }).then((result) => {
          /* Read more about isConfirmed, isDenied below */
          if (result.isConfirmed) {
            // Create an anchor element for downloading the file
            const a = document.createElement("a");
            a.href = directoryPath;
            a.download = storeFile;
            // Trigger a click event on the anchor element to initiate the download
            a.click();
          } else if (result.isDismissed) {
            Swal.fire({
              title: "You have not downloaded!",
              text: `Your file is stored in this ${directoryPath}`,
              icon: "info",
              confirmButtonText: "Ok",
            });
          }
        });
      }
    }
  }
});
