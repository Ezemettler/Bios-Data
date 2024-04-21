function OrdenFunciones() {
  Ventas_TotalACobrar();
  Ventas_CostoVenta();
  Ventas_Mes();
  Egresos_Mes();
  Resultados();
}
  
function Ventas_TotalACobrar() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var lastRow = ventasSheet.getLastRow();
  
  // Obtiene el rango de las columnas E y F para las filas nuevas
  var precios = ventasSheet.getRange("E2:E" + lastRow).getValues();
  var cantidades = ventasSheet.getRange("F2:F" + lastRow).getValues();
  var totalACobrar = ventasSheet.getRange("V2:V" + lastRow).getValues();
  
  // Calcula el total solo para las filas nuevas donde el campo "Total a Cobrar" está vacío
  for (var i = 0; i < precios.length; i++) {
    var total = totalACobrar[i][0];
    if (total === "") { // Comprueba si el campo "Total a Cobrar" está vacío
      var precio = precios[i][0];
      var cantidad = cantidades[i][0];
      var nuevoTotal = precio * cantidad;
      ventasSheet.getRange("V" + (i + 2)).setValue(nuevoTotal); // Establece el nuevo total en la columna V
    }
  }
}
  
function Ventas_CostoVenta() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var lastRow = ventasSheet.getLastRow();

  // Obteniendo los rangos de las columnas relevantes
  var comisionRange = ventasSheet.getRange("I2:I" + lastRow);
  var impuestosRange = ventasSheet.getRange("J2:J" + lastRow);
  var costoEnvioRange = ventasSheet.getRange("M2:M" + lastRow);
  var costoVentaRange = ventasSheet.getRange("W2:W" + lastRow);

  // Obteniendo los valores de las columnas relevantes
  var comisionValues = comisionRange.getValues();
  var impuestosValues = impuestosRange.getValues();
  var costoEnvioValues = costoEnvioRange.getValues();

  // Calculando el costo de venta para cada fila
  var costoVentaValues = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var costoVenta = comisionValues[i][0] + impuestosValues[i][0] + costoEnvioValues[i][0];
    costoVentaValues.push([costoVenta]);
  }

  // Actualizando los valores en la columna "Costo de venta"
  costoVentaRange.setValues(costoVentaValues);
}

function Ventas_Mes() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var lastRow = ventasSheet.getLastRow();

  // Añadir encabezado de columna "Mes" en la columna X
  ventasSheet.getRange("X1").setValue("Mes");

  // Obtener rangos de fechas y de la columna "Mes"
  var fechasRange = ventasSheet.getRange("B2:B" + lastRow);
  var mesRange = ventasSheet.getRange("X2:X" + lastRow);

  // Obtener los valores de las fechas
  var fechasValues = fechasRange.getValues();

  // Array para almacenar los valores de la columna "Mes"
  var mesValues = [];

  // Iterar sobre las fechas y crear el valor de "Mes" en formato YYYY.MM
  for (var i = 0; i < fechasValues.length; i++) {
    var fecha = fechasValues[i][0];
    if (fecha !== "") {
      var mes = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy.MM");
      mesValues.push([mes]);
    } else {
      mesValues.push([""]); // Si no hay fecha, dejar el valor de la columna "Mes" vacío
    }
  }

  // Escribir los valores de "Mes" en la hoja
  mesRange.setValues(mesValues);
}


function Egresos_Mes() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var egresosSheet = spreadsheet.getSheetByName("Egresos");
  var lastRow = egresosSheet.getLastRow();

  // Añadir encabezado de columna "Mes" en la columna G
  egresosSheet.getRange("G1").setValue("Mes");

  // Obtener rangos de fechas y de la columna "Mes"
  var fechasRange = egresosSheet.getRange("B2:B" + lastRow);
  var mesRange = egresosSheet.getRange("G2:G" + lastRow);

  // Obtener los valores de las fechas
  var fechasValues = fechasRange.getValues();

  // Array para almacenar los valores de la columna "Mes"
  var mesValues = [];

  // Iterar sobre las fechas y crear el valor de "Mes" en formato YYYY.MM
  for (var i = 0; i < fechasValues.length; i++) {
    var fecha = fechasValues[i][0];
    if (fecha !== "") {
      var mes = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy.MM");
      mesValues.push([mes]);
    } else {
      mesValues.push([""]); // Si no hay fecha, dejar el valor de la columna "Mes" vacío
    }
  }

  // Escribir los valores de "Mes" en la hoja
  mesRange.setValues(mesValues);
}

function Resultados() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var resultadosSheet = spreadsheet.getSheetByName("Resultados");
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var egresosSheet = spreadsheet.getSheetByName("Egresos");
  var lastRow = resultadosSheet.getLastRow();
  
  // Agregar encabezados de columna si es la primera vez que se ejecuta el script
  if (lastRow == 0) {
    resultadosSheet.getRange("A1").setValue("Mes");
    resultadosSheet.getRange("B1").setValue("Ingresos");
    resultadosSheet.getRange("C1").setValue("Egresos");
    resultadosSheet.getRange("D1").setValue("Materias primas");
  }
  
  // Limpiar valores anteriores
  resultadosSheet.getRange("B2:D" + lastRow).clear();
  
  // Obtener datos de ventas y egresos
  var ventasData = ventasSheet.getDataRange().getValues();
  var egresosData = egresosSheet.getDataRange().getValues();
  
  // Calcular ingresos y egresos para cada mes
  var resultados = new Map();
  for (var i = 1; i < ventasData.length; i++) {
    var fechaVenta = new Date(ventasData[i][1]); // Considerando que la fecha de venta está en la segunda columna de la hoja "Ventas"
    var mes = fechaVenta.getFullYear() + "." + (fechaVenta.getMonth() + 1);
    var totalVenta = Number(ventasData[i][21]); // Considerando que el total de la venta está en la columna 21 de la hoja "Ventas"
    if (!isNaN(totalVenta)) {
      if (!resultados.has(mes)) resultados.set(mes, { ingresos: 0, egresos: 0, materiasPrimas: 0 });
      resultados.get(mes).ingresos += totalVenta;
    }
  }
  for (var i = 1; i < egresosData.length; i++) {
    var fechaEgreso = new Date(egresosData[i][1]); // Considerando que la fecha de pago está en la segunda columna de la hoja "Egresos"
    var mes = fechaEgreso.getFullYear() + "." + (fechaEgreso.getMonth() + 1);
    var categoria = egresosData[i][2]; // Considerando que la categoría está en la tercera columna de la hoja "Egresos"
    var totalEgreso = Number(egresosData[i][4]); // Considerando que el total del egreso está en la columna 4 de la hoja "Egresos"
    if (!isNaN(totalEgreso)) {
      if (!resultados.has(mes)) resultados.set(mes, { ingresos: 0, egresos: 0, materiasPrimas: 0 });
      resultados.get(mes).egresos += totalEgreso;
      if (categoria === "Materias primas") {
        resultados.get(mes).materiasPrimas += totalEgreso;
      }
    }
  }
  
  // Sumar costo de venta de la hoja "Ventas" a los egresos
  for (var i = 1; i < ventasData.length; i++) {
    var fechaVenta = new Date(ventasData[i][1]);
    var mes = fechaVenta.getFullYear() + "." + (fechaVenta.getMonth() + 1);
    var costoVenta = Number(ventasData[i][22]); // Considerando que el costo de venta está en la columna 22 de la hoja "Ventas"
    if (!isNaN(costoVenta)) {
      if (!resultados.has(mes)) resultados.set(mes, { ingresos: 0, egresos: 0, materiasPrimas: 0 });
      resultados.get(mes).egresos += costoVenta;
    }
  }
  
  // Escribir los resultados en la hoja "Resultados" ordenando por mes
  var row = 2;
  [...resultados.keys()].sort().forEach(mes => {
    var ingresos = resultados.get(mes).ingresos;
    var egresos = resultados.get(mes).egresos;
    var materiasPrimas = resultados.get(mes).materiasPrimas;
    resultadosSheet.getRange("A" + row).setValue(mes);
    resultadosSheet.getRange("B" + row).setValue(ingresos);
    resultadosSheet.getRange("C" + row).setValue(egresos);
    resultadosSheet.getRange("D" + row).setValue(materiasPrimas);
    row++;
  });
}
