const fs = require("fs");
const ExcelJS = require("exceljs");

let archivoRAW = fs.readFileSync(__dirname + "/pruebas/isaijsonexample.json");
let archivoJSON = JSON.parse(archivoRAW);
const workbook = new ExcelJS.Workbook();
workbook.creator = "IsaiT";
workbook.created = new Date();
workbook.calcProperties.fullCalcOnLoad = true;
let sheet;
let encabezados = [];

function jsonToExcel(
  objeto,
  padre,
  abuelo,
  level,
  crearPagina,
  bisAbuelo,
  reiniciarLevel,
  filaInser
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

      let esPrinOsec = abuelo.search("-");
      if (esPrinOsec === -1) {
        //Es hoja principal
        let fila = sheet.getRow(1);

        let celda;
        encabezados = fila.values;
        fila = sheet.getRow(1);
        if (fila === null || !fila.values || !fila.values.length) {
          encabezados.push("");
          encabezados.push(collection);
        } else {
          if (!encabezados.includes(collection)) {
            encabezados.push(collection);
          }
        }

        fila.values = encabezados;
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
        fila = sheet.getRow(1);

        let last = sheet.getColumn(1);
        let ultimaFila;
        last.eachCell(function (cel, rowNum) {
          ultimaFila = rowNum;
        });

        let filaTemp;
        console.log(sheet.getRow(1).values);
        console.log(sheet.getRow(2).values);
        console.log(sheet.getRow(3).values);
        console.log(sheet.getRow(4).values);

        if (ultimaFila === 1) {
          ultimaFila = ultimaFila + 1;
          filaTemp = sheet.getRow(ultimaFila).values;
        } else {
          filaTemp = sheet.getRow(ultimaFila).values[1];
        }
        console.log(filaTemp);
        if (filaTemp !== padre) {
          ultimaFila = ultimaFila + 1;
        }

        fila = sheet.getRow(1);
        //Insertar id
        cell = sheet.getCell(ultimaFila, 1);
        cell.value = padre;
        console.log("El ID: " + collection + " va en:" + cell.address);

        fila = sheet.getRow(1);
        //Insertar celda
        cell = sheet.getCell(ultimaFila, numeroCol);
        cell.value = objeto[collection];

        fila = sheet.getRow(1);

        console.log(
          "La collection: " + objeto[collection] + " va en:" + cell.address
        );

        fila = sheet.getRow(1);
        console.log("________________________________________________________");

        cell = null;
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
    } else {
      if (level === 1) {
        jsonToExcel(
          objeto[collection],
          collection,
          collection,
          level,
          "crear",
          bisAbuelo,
          "no",
          filaInser
        );
      } else if (level === 2) {
        jsonToExcel(
          objeto[collection],
          collection,
          abuelo,
          level,
          "seleccionar",
          bisAbuelo,
          "no",
          filaInser + 1
        );
      } else if (level === 3) {
        jsonToExcel(
          objeto[collection],
          collection,
          abuelo + "-" + collection,
          level,
          "crear",
          padre,
          "si",
          (filaInser = 1)
        );
      }
    }
  }
  console.log("TERMINAMOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOS");
  reiniciarLevel = "no";
}

jsonToExcel(archivoJSON, "", "", 0, "", "", "no", 1);

workbook.xlsx.writeFile(__dirname + "/public/Generado/test.xlsx");
