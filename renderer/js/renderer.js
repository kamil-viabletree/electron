const XLSX = require("xlsx");
const Swal = require("sweetalert2");
const path = require("path");
const fs = require("fs");

document.addEventListener("DOMContentLoaded", function () {
  function createDirectoryIfNotExists(directoryPath) {
    if (!fs.existsSync(directoryPath)) {
      fs.mkdirSync(directoryPath, { recursive: true });
    }
  }

  let directoryPath = "";
  let dir = "";
  let downloadDir = "";
  if (process.platform == "linux") {
    const components = __dirname.split("/");
    if (components.length >= 2) {
      downloadDir = components.slice(0, 3).join("/");
      directoryPath = path.join(downloadDir, "Konica");
      createDirectoryIfNotExists(directoryPath);
    } else {
      console.log("String does not contain enough components.");
    }
  } else {
    downloadDir = __dirname.match(/^[A-Z]:/i)[0];
    directoryPath = path.join(downloadDir, "Konica");
    dir =
      "E:/Dropbox (Stupell Industries)/Shared Picklists/ReRunPrep Picklists";
    createDirectoryIfNotExists(directoryPath);
  }

  document.getElementById("myForm").addEventListener("submit", function (e) {
    e.preventDefault();
    const seq_one = document.getElementById("sequence-one").value;
    const seq_two = document.getElementById("sequence-two").value;
    const seq_three = document.getElementById("sequence-three").value;
    const sequencrArr = [seq_one, seq_two, seq_three];
    const excel_file = document.getElementById("excel-file");

    if (
      seq_one.trim() == "" ||
      seq_two.trim() == "" ||
      seq_three.trim() == ""
    ) {
      Swal.fire({
        title: "Error!",
        text: "All sequences should be pre filled!",
        icon: "error",
        confirmButtonText: "OK",
      });
    } else {
      if (excel_file.files.length > 0) {
        const selectedFile = excel_file.files[0];
        const selectedFileName = selectedFile.name;
        const pattern = /statusReport_([0-9]{2}-[0-9]{2} [A-Za-z]{3})/;
        const match = pattern.exec(selectedFileName);

        if (match) {
          // const fileName = match[1];
          const fileName = selectedFileName.replace("statusReport", "RR");
          const reader = new FileReader();

          reader.onload = function (e) {
            const data = e.target.result;

            // const workbook = XLSX.read(data, { type: "binary" });
            processFile(data, sequencrArr, fileName, dir, downloadDir);
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
    }
  });

  function processFile(data, sequencrArr, fileName, dir, downloadDir) {
    let storeFile = fileName;

    const workbook = XLSX.read(data, { type: "binary" });

    const sheetName = workbook.SheetNames[2];
    if (sheetName) {
      const sheet = workbook.Sheets[sheetName];
      if (sheet) {
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        let count = 3;
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
                if (count <= 2) {
                  count++;
                }
              }
            }
          }
        }
        if (printFiles.length > 0 || missingFiles.length > 0) {
          createNewSheet(printFiles, missingFiles, storeFile, dir, downloadDir);
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
        title: "Info!",
        text: "No data found",
        icon: "info",
        confirmButtonText: "OK",
      });
    }
  }

  function getCurrentDateTimeString() {
    var currentDate = new Date();

    var year = currentDate.getFullYear();
    var month = (currentDate.getMonth() + 1).toString().padStart(2, "0"); // Adding 1 to month to make it 1-12 and padding with '0'.
    var day = currentDate.getDate().toString().padStart(2, "0"); // Padding with '0'.
    var hours = currentDate.getHours().toString().padStart(2, "0");
    var minutes = currentDate.getMinutes().toString().padStart(2, "0");
    var seconds = currentDate.getSeconds().toString().padStart(2, "0");

    // Create the formatted date and time string
    var dateTimeString =
      year +
      "-" +
      month +
      "-" +
      day +
      "_" +
      hours +
      ":" +
      minutes +
      ":" +
      seconds;

    return dateTimeString;
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

  function createNewSheet(
    printFiles,
    missingFiles,
    storeFile,
    dir,
    downloadDir
  ) {
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
    let directoryPath = "";
    if (fs.existsSync(dir)) {
      // If directory exists
      directoryPath = path.join(dir, `${storeFile}`);
    } else {
      directoryPath = path.join(downloadDir, `Konica/${storeFile}`);
    }

    if (isFileOpen(directoryPath)) {
      s;
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
