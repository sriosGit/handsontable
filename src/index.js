import Handsontable from "handsontable";
import "handsontable/dist/handsontable.min.css";
import "pikaday/css/pikaday.css";

import {
  alignHeaders,
  addClassesToRows,
  changeCheckboxCell,
} from "./hooksCallbacks";

import Dropzone from "dropzone";
import { read, utils, write } from "xlsx";
import axios from "axios";

const example = document.getElementById("handsontable");
const output = document.querySelector('#output');

console.log("hola");

let myDropzone = new Dropzone("#my-great-dropzone");
const headers = ["First name", "Last name", "Email"];

const emailValidator =  /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const hotTable = new Handsontable(example, {
  colHeaders: headers,
  columns: [
    { data: 0, type: "text" },
    { data: 1, type: "text" },
    { data: 2, type: "text", validator: emailValidator, allowInvalid: true  },
    /*
    {
      data: 4,
      type: "date",
      allowInvalid: false,
    },
    { data: 5, type: "text" },
    {
      data: 6,
      type: "checkbox",
      className: "htCenter",
    },
    {
      data: 7,
      type: "numeric",
    },
    {
      data: 8,
      renderer: progressBarRenderer,
      readOnly: true,
      className: "htMiddle",
    },
    {
      data: 9,
      renderer: starRenderer,
      readOnly: true,
      className: "star htCenter",
    },
    */
  ],
  dropdownMenu: true,
  hiddenColumns: {
    indicators: true,
  },
  contextMenu: true,
  multiColumnSorting: true,
  filters: true,
  rowHeaders: true,
  manualRowMove: true,
  afterGetColHeader: alignHeaders,
  afterOnCellMouseDown: changeCheckboxCell,
  beforeRenderer: addClassesToRows,
  licenseKey: "non-commercial-and-evaluation",
  beforeChange(changes, source) {
    for (let i = changes.length - 1; i >= 0; i--) {
      // capitalise first letter in column 1 and 2

      if ((changes[i][1] === 0 || changes[i][1] === 1)) {
        if(changes[i][3] !== null){
          changes[i][3] = changes[i][3].charAt(0).toUpperCase() + changes[i][3].slice(1);
        }
      }
    }
  },
  afterChange(changes, source) {
    if (source !== 'loadData') {
      output.innerText = JSON.stringify(changes);
    }
  },
});

// DROPZONE

myDropzone.on("addedfile", (file) => {
  const reader = new FileReader();
  reader.readAsArrayBuffer(file);
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = read(data, { type: "array" });

    // Convertir la primera hoja del Excel a JSON
    const jsonData = utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    var formatData = jsonData.map(function (row) {
      return Object.values(row);
    });
    console.log(formatData);
    hotTable.loadData(formatData);
  };
});

// EXPORT DATA

const exportPlugin = hotTable.getPlugin("exportFile");
var downloadBtn = document.getElementById("download");
downloadBtn.addEventListener("click", () => {
  exportPlugin.downloadFile("csv", {
    bom: false,
    columnDelimiter: ",",
    columnHeaders: false,
    exportHiddenColumns: true,
    exportHiddenRows: true,
    // fileExtension: "xlsx",
    filename: "Handsontable-CSV-file_[YYYY]-[MM]-[DD]",
    // mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    rowDelimiter: "\r\n",
    rowHeaders: true,
  });
});

// Manejo del evento clic para el botón de enviar JSON
var sendDataButton = document.getElementById("sendJson");
sendDataButton.addEventListener("click", function () {
  const sendDataObject = hotTable.getData();
  sendDataObject.unshift(headers);
  console.log(sendDataObject);

  axios
    .post("send_as_array", { sendDataObject })
    .then(function (response) {
      console.log("Datos enviados con éxito:", response.data);
    })
    .catch(function (error) {
      console.error("Error al enviar datos:", error);
    });
});

function createExcel() {
  var excelData = hotTable.getData(); // Obtener la data de Handsontable
  excelData.unshift(headers); // Agregar encabezados al inicio

  // Crear un nuevo libro de trabajo de Excel
  var wb = utils.book_new();
  var ws = utils.aoa_to_sheet(excelData); // Convertir el arreglo de arreglos a hoja de cálculo

  utils.book_append_sheet(wb, ws, "Hoja1"); // Agregar la hoja al libro de trabajo

  // Convertir el libro de trabajo a un archivo binario
  var excelBuffer = write(wb, { bookType: "xlsx", type: "array" });

  // Crear un Blob a partir del buffer
  var blob = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return blob;
}

// Manejo del evento clic para el botón de enviar Excel
var sendExcelButton = document.getElementById("sendExcel");
sendExcelButton.addEventListener("click", function () {
  // Crear un objeto FormData para enviar el archivo a través de Axios
  const blob = createExcel();
  var formData = new FormData();
  formData.append("file", blob, "data.xlsx"); // 'file' es el nombre del campo en el formulario

  // Enviar el archivo al endpoint utilizando Axios
  axios
    .post("send_as_excel", formData, {
      headers: {
        "Content-Type": "multipart/form-data",
      },
    })
    .then(function (response) {
      console.log("Archivo Excel enviado con éxito:", response.data);
    })
    .catch(function (error) {
      console.error("Error al enviar el archivo Excel:", error);
    });
});

// Manejo del evento clic para el botón de enviar Excel
var downloadExcelButton = document.getElementById("downloadExcel");
downloadExcelButton.addEventListener("click", function () {
  // Crear un objeto FormData para enviar el archivo a través de Axios
  const blob = createExcel();
  var fileName = "data.xlsx"; // Nombre del archivo
  var downloadLink = document.createElement("a");

  downloadLink.href = window.URL.createObjectURL(blob);
  downloadLink.download = fileName;
  downloadLink.click();

  // Enviar el archivo al endpoint utilizando Axios
  axios
    .post("TU_ENDPOINT_EXCEL", formData, {
      headers: {
        "Content-Type": "multipart/form-data",
      },
    })
    .then(function (response) {
      console.log("Archivo Excel enviado con éxito:", response.data);
    })
    .catch(function (error) {
      console.error("Error al enviar el archivo Excel:", error);
    });
});
