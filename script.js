// Cached DOM elements
const table = document.querySelector(".sheet-body");
const rowsInput = document.querySelector(".rows");
const columnsInput = document.querySelector(".columns");
let tableExists = false;

// Utility function to show alerts
const showAlert = (title, text, icon) => {
  Swal.fire({ title, text, icon });
};

// Function to generate the table
const generateTable = () => {
  const rowsNumber = parseInt(rowsInput.value);
  const columnsNumber = parseInt(columnsInput.value);

  if (isNaN(rowsNumber) || isNaN(columnsNumber)) {
    showAlert("Error!", "Please enter valid numbers!", "error");
    return;
  }

  if (rowsNumber < 1 || columnsNumber < 1) {
    showAlert("Error!", "Please enter positive numbers!", "error");
    return;
  }

  table.innerHTML = "";
  for (let i = 0; i < rowsNumber; i++) {
    let tableRow = "<tr>";
    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable></td>`;
    }
    tableRow += "</tr>";
    table.innerHTML += tableRow;
  }

  tableExists = true;
  showAlert("Good job!", "Table generated successfully!", "success");
};

// Function to export the table to Excel
const exportToExcel = (type, filename, download) => {
  if (!tableExists) {
    showAlert("Error!", "Please generate a table first!", "error");
    return;
  }

  const workbook = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
  if (download) {
    XLSX.write(workbook, { bookType: type, bookSST: true, type: "base64" });
  } else {
    XLSX.writeFile(workbook, filename || `MyNewSheet.${type || "xlsx"}`);
  }

  showAlert("Success!", "Table exported successfully.", "success");
  tableExists = false;
};

// Event listeners for buttons
const generateButton = document.getElementById("generateButton");
const exportButton = document.getElementById("exportButton");

generateButton.addEventListener("click", generateTable);
exportButton.addEventListener("click", () => exportToExcel("xlsx"));

