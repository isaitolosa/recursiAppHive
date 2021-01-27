const fs = require("fs");
const ExcelJS = require("exceljs");

let archivoRAW = fs.readFileSync(__dirname + "/pruebas/isaijsonexample.json");
let archivoJSON = JSON.parse(archivoRAW);
const workbook = new ExcelJS.Workbook();
workbook.creator = "IsaiT";
workbook.created = new Date();
workbook.calcProperties.fullCalcOnLoad = true;
let sheet;

function jsonToExcel(
  objeto,
  padre,
  abuelo,
  level,
  crearPagina,
  bisAbuelo,
  reiniciarLevel
) {
  console.log("Se llamo la funcion_____________________________");
  if (reiniciarLevel === "si") {
    level = 2;
  } else {
    level = level + 1;
  }

  //Este ciclo va a pasar por los nodos del mas grande al mas pequeño
  if (crearPagina === "crear") {
    if (workbook.getWorksheet(abuelo) === undefined) {
      sheet = workbook.addWorksheet(abuelo);
    } else {
      sheet = workbook.getWorksheet(abuelo);
    }
  } else if (crearPagina === "seleccionar") {
    sheet = workbook.getWorksheet(abuelo);
  }

  let filaAinsertar = 2;
  for (var collection in objeto) {
    if (typeof objeto[collection] !== "object") {
      sheet = workbook.getWorksheet(abuelo);

      console.log(
        "Los valores son: " +
          collection +
          ": " +
          objeto[collection] +
          ". En la página: " +
          abuelo +
          "."
      );
      //sheet.addRow([3, "Sam", new Date()]);

      let esPrinOsec = abuelo.search(":");
      if (esPrinOsec === -1) {
        //Es hoja principal
        let fila = sheet.getRow(1);
        let encabezados = [];
        let celda;
        encabezados = fila.values;
        console.log(encabezados);
        if (fila == null || !fila.values || !fila.values.length) {
          celda = sheet.getCell(1, 2);
          celda.value = collection;
          encabezados.push("");
          encabezados.push(collection);
        } else {
          if (!encabezados.includes(collection)) {
            encabezados.push(collection);
          }
        }

        sheet.insertRow(1, encabezados);
        fila = sheet.getRow(1);
        console.log(fila.values);
        //console.log(fila.values);

        //Buscar nombre de columna en la fila y sacar su posicion

        let numeroCol;
        fila.eachCell(function (cel, rowNum) {
          if (cel.text === collection) {
            numeroCol = rowNum;
          }
        });
        console.log(numeroCol);

        //Insertar id
        cell = sheet.getCell(filaAinsertar, 1);
        cell.value = padre;
        //Insertar celda
        cell = sheet.getCell(filaAinsertar, numeroCol);
        cell.value = objeto[collection];
      } else {
        //Es hoja anidada
        let fila = sheet.getRow(1);
        let encabezados = [];
        let celda;
        if (fila == null || !fila.values || !fila.values.length) {
          celda = sheet.getCell(1, 3);
          celda.value = collection;
        } else {
          for (let i = 1; i < fila.values.length; i++) {
            let cell = fila.getCell(i).text;
            encabezados.push(cell);
          }
        }
      }

      console.log();
    } else {
      if (level === 1) {
        jsonToExcel(
          objeto[collection],
          collection,
          collection,
          level,
          "crear",
          bisAbuelo,
          "no"
        );
      } else if (level === 2) {
        jsonToExcel(
          objeto[collection],
          collection,
          abuelo,
          level,
          "seleccionar",
          bisAbuelo,
          "no"
        );
      } else if (level === 3) {
        jsonToExcel(
          objeto[collection],
          collection,
          abuelo + "-" + collection,
          level,
          "crear",
          padre,
          "si"
        );
      }
    }
  }
  console.log("TERMINAMOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOS");
  reiniciarLevel = "no";
}

workbook.xlsx.writeFile(__dirname + "public/Generado/test.xlsx");

jsonToExcel(archivoJSON, "", "", 0, "", "", "no");
