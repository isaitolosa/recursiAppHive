const fs = require("fs");
const ExcelJS = require("exceljs");
const excelToJson = require("convert-excel-to-json");
//import { JsonDB } from "node-json-db";
//import { Config } from "node-json-db/dist/lib/JsonDBConfig";
const { JsonDB } = require("node-json-db");
const { Config } = require("node-json-db/dist/lib/JsonDBConfig");

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

        //console.log(fila.values);

        //Buscar nombre de columna en la fila y sacar su posicion

        let numeroCol;
        fila.eachCell(function (cel, rowNum) {
          if (cel.text === collection) {
            numeroCol = rowNum;
          }
        });
        fila = sheet.getRow(1);

        let last = sheet.getColumn(1);
        let ultimaFila;
        last.eachCell(function (cel, rowNum) {
          ultimaFila = rowNum;
        });

        let filaTemp;

        if (ultimaFila === 1) {
          ultimaFila = ultimaFila + 1;
          filaTemp = sheet.getRow(ultimaFila).values;
        } else {
          filaTemp = sheet.getRow(ultimaFila).values[1];
        }
        if (filaTemp.length === 0) {
        } else if (filaTemp !== padre) {
          ultimaFila = ultimaFila + 1;
        }

        fila = sheet.getRow(1);
        //Insertar id
        cell = sheet.getCell(ultimaFila, 1);
        cell.value = padre;

        fila = sheet.getRow(1);
        //Insertar celda
        cell = sheet.getCell(ultimaFila, numeroCol);
        cell.value = objeto[collection];

        fila = sheet.getRow(1);

        fila = sheet.getRow(1);

        cell = null;
      } else {
        //Es hoja anidada
        let fila = sheet.getRow(1);

        let celda;
        encabezados = fila.values;
        fila = sheet.getRow(1);
        if (fila === null || !fila.values || !fila.values.length) {
          encabezados.push("");
          encabezados.push("");
          encabezados.push(collection);
        } else {
          if (!encabezados.includes(collection)) {
            encabezados.push(collection);
          }
        }

        fila.values = encabezados;
        fila = sheet.getRow(1);

        //Buscar nombre de columna en la fila y sacar su posicion

        let numeroCol;
        fila.eachCell(function (cel, rowNum) {
          if (cel.text === collection) {
            numeroCol = rowNum;
          }
        });
        fila = sheet.getRow(1);

        let last = sheet.getColumn(1);
        let ultimaFila;
        last.eachCell(function (cel, rowNum) {
          ultimaFila = rowNum;
        });

        let filaTemp;

        if (ultimaFila === 1) {
          ultimaFila = ultimaFila + 1;
          filaTemp = sheet.getRow(ultimaFila).values;
        } else {
          filaTemp = sheet.getRow(ultimaFila).values[2];
        }
        if (filaTemp.length === 0) {
        } else if (filaTemp !== padre) {
          ultimaFila = ultimaFila + 1;
        }

        fila = sheet.getRow(1);
        //Insertar ids
        cell = sheet.getCell(ultimaFila, 1);
        cell.value = bisAbuelo;
        cell = sheet.getCell(ultimaFila, 2);
        cell.value = padre;

        fila = sheet.getRow(1);
        //Insertar celda
        cell = sheet.getCell(ultimaFila, numeroCol);
        cell.value = objeto[collection];

        fila = sheet.getRow(1);

        fila = sheet.getRow(1);

        cell = null;
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
  reiniciarLevel = "no";
}

//jsonToExcel(archivoJSON, "", "", 0, "", "", "no", 1);

//workbook.xlsx.writeFile(__dirname + "/public/Generado/test.xlsx");

var db = new JsonDB(new Config("miBD", true, false, "/"));

function ExcelTojson() {
  const result = excelToJson({
    sourceFile: __dirname + "/public/Generado/test.xlsx",
  });

  function insertaEnbd(path, objeto) {
    //console.log(path);
    //console.log(objeto);
  }

  function funcionRecursi(path, getObjeto) {
    let arreglo = [];
    let paginaAnterior = "asdf#$%&nada....";

    for (var pag in getObjeto) {
      //console.log(pag.search(paginaAnterior));
      for (var nombrePag in getObjeto[pag]) {
        console.log(nombrePag.search(paginaAnterior));
        console.log(getObjeto[pag][nombrePag]);
      }
    }
  }

  //____________________________________________________________________________
  let collectionAnterior = "";
  let objetoAmandar = [];
  for (var pagina in result) {
    //pagina da el nombre de la página en excel
    if (collectionAnterior === "") {
      collectionAnterior = pagina;
    }

    let separaCollections = pagina.split("-");
    if (separaCollections.length === 1) {
      if (collectionAnterior !== pagina) {
        //no es igual, hacer el reset de variables, analizar cuando la ultima página es principal(collection)
        //1: llama la función para insertar lo que esta en el objetoAmandar
        funcionRecursi("", objetoAmandar);
        //2: inserta la pagina(collection) actual a la bd
        insertaEnbd(pagina, result[pagina]);
        //3: borrar la información que está en objetoAmandar = []
        objetoAmandar = [];
      } else {
        //insertar primera pagina en db
        insertaEnbd(pagina, result[pagina]);
      }
    } else {
      //console.log(pagina + ":{" + result[pagina] + "}");
      let paraVer = result[pagina];
      //console.log();
      let cadena = '{"' + pagina + '":' + JSON.stringify(result[pagina]) + "}";
      let objeto = JSON.parse(cadena);
      objetoAmandar.push(objeto);
      //console.log(objetoAmandar);
    }
    collectionAnterior = pagina;
  }

  //Checamos si el objetoAmandar no esta vacío, llamar a la funcion una ultima vez
  if (Object.keys(objetoAmandar).length !== 0) {
    funcionRecursi(collectionAnterior, objetoAmandar);
    objetoAmandar = [];
  }

  //__________________________________________________________________________

  //console.log(result);
  //console.log(typeof Object.keys(result));
  ///console.log(Object.keys(result).length);

  //db.push("/" + pagina, result[pagina], false);
  //db.delete("/");
  var data = db.getData("/");
  //console.log(data);
}

ExcelTojson();
