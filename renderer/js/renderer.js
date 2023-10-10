const XLSX = require("xlsx");
const Swal = require("sweetalert2");

document.addEventListener("DOMContentLoaded", function () {
  document.getElementById("myForm").addEventListener("submit", function (e) {
    e.preventDefault();
    const seq_one = document.getElementById("sequence-one").value;
    const seq_two = document.getElementById("sequence-two").value;
    const seq_three = document.getElementById("sequence-three").value;
    const sequencrArr = [seq_one, seq_two, seq_three];
    const excel_file = document.getElementById("excel-file");

    if (excel_file.files.length > 0) {
      const selectedFile = excel_file.files[0];
      // const selectedFileName = selectedFile.name;
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = e.target.result;

        // const workbook = XLSX.read(data, { type: "binary" });
        processFile(data, sequencrArr);
      };
      reader.readAsBinaryString(selectedFile);
      document.getElementById("myForm").reset();
      Swal.fire({
        title: "Done!",
        text: "Sheet processed successfully",
        icon: "success",
        confirmButtonText: "Cool",
      });
    } else {
      Swal.fire({
        title: "Error!",
        text: "Konica Print sheet not found!",
        icon: "error",
        confirmButtonText: "Cool",
      });
    }
  });

  function processFile(data, sequencrArr) {
    let storeFile = "RR_10-04 Wed Konic.xlsx";

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
          createNewSheet(printFiles, missingFiles, storeFile);
        } else {
          // error
          Swal.fire({
            title: "Info!",
            text: "No remmainings were found",
            icon: "info",
            confirmButtonText: "Cool",
          });
        }
      } else {
        Swal.fire({
          title: "Info!",
          text: "No data found",
          icon: "info",
          confirmButtonText: "Cool",
        });
      }
    } else {
      Swal.fire({
        title: "Info!",
        text: "No data found",
        icon: "info",
        confirmButtonText: "Cool",
      });
    }
  }

  /**
   * New Sheet creating
   */

  function createNewSheet(printFiles, missingFiles, storeFile) {
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
    XLSX.writeFile(newWorkbook, storeFile);
  }
});
