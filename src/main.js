import "./style.css";
import Papa from "papaparse";
import * as XLSX from "xlsx";

document.querySelector("#app").innerHTML = `
  <div>
    <h1>ii Access Report Transformer</h1>
    <div id="drop-area" class="drop-area">
      <p>Drag & drop a CSV file here, or <button id="browse-btn">browse</button></p>
      <input type="file" id="csv-file" accept=".csv" />
    </div>

    <div id="status"></div>
  </div>
`;

const fileInput = document.querySelector("#csv-file");
const status = document.querySelector("#status");
const dropArea = document.querySelector("#drop-area");
const browseBtn = document.querySelector("#browse-btn");

let rawData = [];

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  parseCSV(file);
});

browseBtn.addEventListener("click", (e) => {
  e.preventDefault();
  fileInput.click();
});

["dragenter", "dragover"].forEach((evt) => {
  dropArea.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropArea.classList.add("dragover");
  });
});

["dragleave", "drop", "dragend"].forEach((evt) => {
  dropArea.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropArea.classList.remove("dragover");
  });
});

dropArea.addEventListener("drop", (e) => {
  const dt = e.dataTransfer;
  if (!dt) return;
  const file = dt.files[0];
  if (!file) return;
  parseCSV(file);
});

function parseCSV(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => {
      rawData = results.data;
      status.innerText = `Loaded ${rawData.length} rows.`;
      processData(rawData);
    },
    error: (err) => {
      status.innerText = `Error parsing file: ${err}`;
    },
  });
}

function processData(data) {
  const feiEmployees = data
    .filter((row) => {
      return row?.Details.includes("FEI");
    })
    .map(({ Number, Reason, Video, ...rest }) => ({
      ...rest,
    }));

  const groupedByDay = Object.groupBy(feiEmployees, (row) =>
    new Date(row.Time).toLocaleDateString("en-GB"),
  );

  Object.entries(groupedByDay).forEach(([date, employees]) => {
    saveAsExcel(
      employees,
      `Forward Emphasis Access Report ${date.replaceAll("/", ".")}.xlsx`,
    );
  });
}

function saveAsExcel(data, filename) {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  // Generates the file and triggers download
  XLSX.writeFile(workbook, filename);
  status.innerText = "File downloaded successfully.";
}
