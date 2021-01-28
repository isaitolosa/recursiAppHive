const express = require("express");
const fs = require("fs");
const ExcelJS = require("exceljs");
const excelToJson = require("convert-excel-to-json");
//import { JsonDB } from "node-json-db";
//import { Config } from "node-json-db/dist/lib/JsonDBConfig";
const { JsonDB } = require("node-json-db");
const { Config } = require("node-json-db/dist/lib/JsonDBConfig");
const { log } = require("console");
const multer = require("multer");
const path = require("path");

const app = express();

// -> Multer Upload Storage
const storage = multer.diskStorage({
  destination: path.join(__dirname, "public/uploads"),
  filename: (req, file, cb) => {
    cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname);
  },
});

app.use(
  multer({
    storage,
    dest: path.join(__dirname, "public/uploads"),
  }).single("uploadfile")
);

let archivoRAW = fs.readFileSync(__dirname + "/pruebas/isaijsonexample.json");
let archivoJSON = JSON.parse(archivoRAW);
const workbook = new ExcelJS.Workbook();
workbook.creator = "IsaiT";
workbook.created = new Date();
workbook.calcProperties.fullCalcOnLoad = true;
let sheet;
let encabezados = [];
var db = new JsonDB(new Config("miBD", true, false, "/"));

app.post("/api/jsonToExcel", (req, res) => {
  let archivoRAW = fs.readFileSync(
    __dirname + "/public/uploads/" + req.file.filename
  );
  let archivoJSON = JSON.parse(archivoRAW);

  jsonToExcel(archivoJSON, "", "", 0, "", "", "no", 1);

  workbook.xlsx.writeFile(__dirname + "/public/Generado/test.xlsx");

  var file = __dirname + "/public/Generado/test.xlsx";
  var filename = path.basename(file);
  console.log(file);
  var mimetype =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  res.setHeader("Content-disposition", "attachment; filename=nuevo.xlsx");
  //res.setHeader("Content-type", mimetype);
  res.contentType = "application/vnd.ms-excel";
  res.download(file, filename);
});

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

app.post("/api/ExcelToJson", (req, res) => {
  const result = excelToJson({
    sourceFile: __dirname + "/public/uploads/" + req.file.filename,
  });
  ExcelTojson(result);
  var data = db.getData("/");
  db.delete("/");
  console.log("Terminamos");
  res.json(data);
});

function ExcelTojson(result) {
  function buscarNodo(paths, laCadena, aEncontrar) {
    for (var num in paths) {
      var llave = Object.keys(paths[num]);
      if (llave.toString() === laCadena) {
        var valores = JSON.parse(Object.values(paths[num]));
        if (aEncontrar === Object.keys(valores).toString()) {
          return Object.values(valores).toString();
        }
      }
    }
  }

  function insertaEnbd(path, objeto) {
    var i = 0;
    var rutaTemp;
    var encabezados;
    let a, b;
    let guardaCeldaAAA;
    let encontrados = [];
    let cadenaBuena;
    for (var numero in objeto) {
      for (var pagina in objeto[numero]) {
        for (var linea in objeto[numero][pagina]) {
          var fila = objeto[numero][pagina][linea];
          if (i === 0) {
            encabezados = objeto[numero][pagina][linea];
            i = i + 1;
          } else {
            //Sección para buscar los encabezados
            for (var individuo in fila) {
              let separaCollections = pagina.split("-");
              if (separaCollections.length === 1) {
                if (individuo === "A") {
                  a = fila[individuo];
                } else {
                  let elEncabezadoIndiv = encabezados[individuo];
                  let elValor = fila[individuo];
                  if (typeof elValor === "string") {
                    elValor = elValor.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elValor = elValor.trim();

                    elEncabezadoIndiv = elEncabezadoIndiv.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elEncabezadoIndiv = elEncabezadoIndiv.trim();
                  }

                  let aux =
                    '{"' +
                    a +
                    '":{"' +
                    elEncabezadoIndiv +
                    '":"' +
                    elValor +
                    '"}}';
                  let jso = JSON.parse(aux);
                  db.push("/" + pagina, jso, false);
                }
              } else if (separaCollections.length === 2) {
                if (individuo === "A") {
                  a = fila[individuo];
                } else if (individuo === "B") {
                  b = fila[individuo];
                } else {
                  let elEncabezadoIndiv = encabezados[individuo];
                  let elValor = fila[individuo];
                  if (typeof elValor === "string") {
                    elValor = elValor.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elValor = elValor.trim();

                    elEncabezadoIndiv = elEncabezadoIndiv.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elEncabezadoIndiv = elEncabezadoIndiv.trim();
                  }
                  let aux =
                    '{"' +
                    a +
                    '":{"' +
                    separaCollections[1] +
                    '":{"' +
                    b +
                    '":{"' +
                    elEncabezadoIndiv +
                    '":"' +
                    elValor +
                    '"}}}}';
                  let jso = JSON.parse(aux);
                  db.push("/" + separaCollections[0], jso, false);
                }
              } else {
                if (individuo === "A") {
                  a = fila[individuo];
                  guardaCeldaAAA = a;
                } else if (individuo === "B") {
                  b = fila[individuo];
                } else {
                  let elEncabezadoIndiv = encabezados[individuo];
                  let elValor = fila[individuo];
                  if (typeof elValor === "string") {
                    elValor = elValor.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elValor = elValor.trim();

                    elEncabezadoIndiv = elEncabezadoIndiv.replace(
                      /(\r\n|\n|\r|\t|[']|["]|\\|\")/gm,
                      ""
                    );
                    elEncabezadoIndiv = elEncabezadoIndiv.trim();
                  }
                  let aux = '{"' + elEncabezadoIndiv + '":"' + elValor + '"}';
                  let jso = JSON.parse(aux);
                  let limite = separaCollections.length - 2;
                  let cadenaBuena = "";
                  let otroLimit = 1;
                  for (
                    let index = limite;
                    index < separaCollections.length;
                    index++
                  ) {
                    let buscarEnPagina = [];
                    //Los dos for sirven para convertir y concatenar el obj en un string
                    for (
                      let z = 0;
                      z < separaCollections.length - otroLimit;
                      z++
                    ) {
                      buscarEnPagina.push(separaCollections[z]);
                    }
                    let laCadena;
                    for (let l = 0; l < buscarEnPagina.length; l++) {
                      if (l === 0) {
                        laCadena = buscarEnPagina[l];
                      } else {
                        laCadena = laCadena + "-" + buscarEnPagina[l];
                      }
                    }
                    if (buscarEnPagina.length === 1) {
                    } else {
                      let aux2 = buscarNodo(path, laCadena, a);
                      if (!(aux2 === undefined)) {
                        encontrados.push(aux2);
                        a = aux2;
                      }
                    }
                    otroLimit++;
                  }
                  //Aqui procedemos a insertar
                  let penultimo =
                    separaCollections[separaCollections.length - 2];
                  let ultimo = separaCollections[separaCollections.length - 1];
                  cadenaBuena =
                    penultimo + "/" + guardaCeldaAAA + "/" + ultimo + "/" + b;
                  let cadenaSumada = "";
                  let encontradosAux = [].concat(encontrados);
                  encontradosAux.reverse();
                  while (encontradosAux.length > 0) {
                    let otroAux = [].concat(encontradosAux);
                    encontradosAux.pop();
                    cadenaSumada =
                      separaCollections[encontradosAux.length] +
                      "/" +
                      otroAux.pop() +
                      "/" +
                      cadenaSumada;
                  }
                  cadenaBuena = cadenaSumada + cadenaBuena;
                  db.push("/" + cadenaBuena, jso, false);
                }
              }
            }
            encontrados = [];
          }
        }
      }
    }
  }

  function funcionRecursi(pathAenviar, getObjeto) {
    let arregloAenviar = [];
    let paginaAnterior = "asdf#$%&nada....";
    let objetoAinsertar = [];
    let paginaActual = "";
    for (var pag in getObjeto) {
      for (var nombrePag in getObjeto[pag]) {
        if (objetoAinsertar.length === 0) {
          let cadena =
            '{"' +
            nombrePag +
            '":' +
            JSON.stringify(getObjeto[pag][nombrePag]) +
            "}";
          let objeto = JSON.parse(cadena);
          objetoAinsertar.push(objeto);
          paginaActual = nombrePag;
        } else {
          let cadena =
            '{"' +
            nombrePag +
            '":' +
            JSON.stringify(getObjeto[pag][nombrePag]) +
            "}";
          let objeto = JSON.parse(cadena);
          arregloAenviar.push(objeto);
        }
      }
      paginaAnterior = nombrePag;
    }

    let separaCollections = paginaActual.split("-");
    if (separaCollections.length > 1) {
      for (var numero in objetoAinsertar) {
        let i = 0;
        let x = 1;
        for (var pagina in objetoAinsertar[numero]) {
          for (var linea in objetoAinsertar[numero][pagina]) {
            var linea = objetoAinsertar[numero][pagina][linea];
            if (i === 0) {
              i = i + 1;
            } else {
              let cadena =
                '{"' +
                paginaActual +
                '":' +
                JSON.stringify('{"' + linea.B + '":"' + linea.A + '"}') +
                "}";
              let objeto = JSON.parse(cadena);
              pathAenviar.push(objeto);
            }
          }
        }
      }
    }
    if (arregloAenviar.length === 0) {
      insertaEnbd(pathAenviar, objetoAinsertar);
    } else {
      insertaEnbd(pathAenviar, objetoAinsertar);
      funcionRecursi(pathAenviar, arregloAenviar);
    }
  }

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
        funcionRecursi([], objetoAmandar);
        //3: borrar la información que está en objetoAmandar = []
        objetoAmandar = [];
        let cadena =
          '{"' + pagina + '":' + JSON.stringify(result[pagina]) + "}";
        let objeto = JSON.parse(cadena);
        objetoAmandar.push(objeto);
      } else {
        //insertar primera pagina en db
        let cadena =
          '{"' + pagina + '":' + JSON.stringify(result[pagina]) + "}";
        let objeto = JSON.parse(cadena);
        objetoAmandar.push(objeto);
        funcionRecursi([], objetoAmandar);
        objetoAmandar = [];
      }
    } else {
      if (pagina.search(collectionAnterior) !== -1) {
        let cadena =
          '{"' + pagina + '":' + JSON.stringify(result[pagina]) + "}";
        let objeto = JSON.parse(cadena);
        objetoAmandar.push(objeto);
      } else {
        funcionRecursi([], objetoAmandar);
        objetoAmandar = [];
        let cadena =
          '{"' + pagina + '":' + JSON.stringify(result[pagina]) + "}";
        let objeto = JSON.parse(cadena);
        objetoAmandar.push(objeto);
      }
    }
    collectionAnterior = pagina;
  }
  //Checamos si el objetoAmandar no esta vacío, llamar a la funcion una ultima vez
  if (Object.keys(objetoAmandar).length !== 0) {
    funcionRecursi([], objetoAmandar);
    objetoAmandar = [];
  }
}

//Settings
app.set("port", 3000);

// Create a Server
app.listen(app.get("port"), () => {
  console.log(`Server on port ${app.get("port")}`);
});
