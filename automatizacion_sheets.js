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
    resultadosSheet.getRange("E1").setValue("Mano de obra");
    resultadosSheet.getRange("F1").setValue("Comisiones");
    resultadosSheet.getRange("G1").setValue("Envíos");
    resultadosSheet.getRange("H1").setValue("Utilidad bruta");
    resultadosSheet.getRange("I1").setValue("Costos financieros");
    resultadosSheet.getRange("J1").setValue("Embalajes");
    resultadosSheet.getRange("K1").setValue("Equipamientos");
    resultadosSheet.getRange("L1").setValue("Fijos");
    resultadosSheet.getRange("M1").setValue("Fletes");
    resultadosSheet.getRange("N1").setValue("Insumos");
    resultadosSheet.getRange("O1").setValue("Publicidad");
    resultadosSheet.getRange("P1").setValue("Servicios");
    resultadosSheet.getRange("Q1").setValue("Utilidad antes de impuestos");
    resultadosSheet.getRange("R1").setValue("Impuestos");
    resultadosSheet.getRange("S1").setValue("Resultado"); 
    resultadosSheet.getRange("T1").setValue("Margen de Ganancia"); // Nuevo encabezado para Margen de Ganancia
  }
  
  // Limpiar valores anteriores
  resultadosSheet.getRange("B2:T" + lastRow).clear(); // Ajuste del rango de limpieza
  
  // Obtener datos de ventas y egresos
  var ventasData = ventasSheet.getDataRange().getValues();
  var egresosData = egresosSheet.getDataRange().getValues();
  
  // Calcular ingresos, egresos, materias primas, mano de obra, comisiones, envíos, utilidad bruta, costos financieros, embalajes, equipamientos, fijos, fletes, insumos, publicidad, servicios, utilidad antes de impuestos, impuestos y resultado para cada mes
  var resultados = new Map();
  for (var i = 1; i < ventasData.length; i++) {
    var fechaVenta = new Date(ventasData[i][1]); 
    var mes = fechaVenta.getFullYear() + "." + (fechaVenta.getMonth() + 1);
    var totalVenta = Number(ventasData[i][21]); 
    var comision = Number(ventasData[i][8]); 
    var costoEnvio = Number(ventasData[i][12]); 
    var impuestosVenta = Number(ventasData[i][9]); // Obtener el valor de impuestos de la venta
    var costoVenta = Number(ventasData[i][22]); // Obtener el costo de venta
    if (!isNaN(totalVenta)) {
      if (!resultados.has(mes)) resultados.set(mes, { ingresos: 0, egresos: 0, materiasPrimas: 0, manoDeObra: 0, comisiones: 0, envios: 0, utilidadBruta: 0, costosFinancieros: 0, embalajes: 0, equipamientos: 0, fijos: 0, fletes: 0, insumos: 0, publicidad: 0, servicios: 0, utilidadAntesImpuestos: 0, impuestos: 0, resultado: 0, margenDeGanancia: 0 }); // Agregar margen de ganancia
      resultados.get(mes).ingresos += totalVenta;
      resultados.get(mes).comisiones += comision;
      resultados.get(mes).envios += costoEnvio;
      resultados.get(mes).impuestos += impuestosVenta; // Agregar impuestos de la venta al total de impuestos del mes
      resultados.get(mes).egresos += costoVenta; // Sumar el costo de venta a los egresos del mes correspondiente
    }
  }

  for (var i = 1; i < egresosData.length; i++) {
      var fechaEgreso = new Date(egresosData[i][1]); 
      var mes = fechaEgreso.getFullYear() + "." + (fechaEgreso.getMonth() + 1);
      var categoria = egresosData[i][2]; 
      var totalEgreso = Number(egresosData[i][4]); 
      var impuestosEgreso = Number(egresosData[i][4]); // Obtener el valor total de impuestos de egreso
      if (!isNaN(totalEgreso)) {
          if (!resultados.has(mes)) resultados.set(mes, { ingresos: 0, egresos: 0, materiasPrimas: 0, manoDeObra: 0, comisiones: 0, envios: 0, utilidadBruta: 0, costosFinancieros: 0, embalajes: 0, equipamientos: 0, fijos: 0, fletes: 0, insumos: 0, publicidad: 0, servicios: 0, utilidadAntesImpuestos: 0, impuestos: 0, resultado: 0 }); 
          resultados.get(mes).egresos += totalEgreso;
          if (categoria === "Materias primas") {
            resultados.get(mes).materiasPrimas += totalEgreso; // Sumar el valor total de materias primas al mes correspondiente
          }
          if (categoria === "Mano de obra") {
            resultados.get(mes).manoDeObra += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Costos financieros") {
            resultados.get(mes).costosFinancieros += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Embalajes") {
            resultados.get(mes).embalajes += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Equipamientos") {
            resultados.get(mes).equipamientos += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Fijos") {
            resultados.get(mes).fijos += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Fletes") {
            resultados.get(mes).fletes += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Insumos") {
            resultados.get(mes).insumos += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Publicidad") {
            resultados.get(mes).publicidad += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Servicios") {
            resultados.get(mes).servicios += totalEgreso; // Sumar el valor total de mano de obra al mes correspondiente
          }
          if (categoria === "Impuestos") {
              resultados.get(mes).impuestos += impuestosEgreso; // Agregar impuestos del egreso al total de impuestos del mes
          }
      }
  }
  
  // Calcular utilidad bruta
  for (var [mes, data] of resultados) {
    data.utilidadBruta = data.ingresos - data.materiasPrimas - data.manoDeObra - data.comisiones - data.envios;
  }
  
  // Calcular utilidad antes de impuestos
  for (var [mes, data] of resultados) {
    data.utilidadAntesImpuestos = data.utilidadBruta - data.costosFinancieros - data.embalajes - data.equipamientos - data.fijos - data.fletes - data.insumos - data.publicidad - data.servicios;
  }
  
  // Calcular resultado
  for (var [mes, data] of resultados) {
    data.resultado = data.utilidadAntesImpuestos - data.impuestos;
  }
  
  // Calcular margen de ganancia
  for (var [mes, data] of resultados) {
    if (data.ingresos === 0) {
        data.margenDeGanancia = 0;
    } else {
        data.margenDeGanancia = data.resultado / data.ingresos;
    }
  }
  
  // Escribir los resultados en la hoja "Resultados" ordenando por mes
  var row = 2;
  [...resultados.keys()].sort().forEach(mes => {
    var ingresos = resultados.get(mes).ingresos;
    var egresos = resultados.get(mes).egresos;
    var materiasPrimas = resultados.get(mes).materiasPrimas;
    var manoDeObra = resultados.get(mes).manoDeObra;
    var comisiones = resultados.get(mes).comisiones;
    var envios = resultados.get(mes).envios;
    var utilidadBruta = resultados.get(mes).utilidadBruta;
    var costosFinancieros = resultados.get(mes).costosFinancieros;
    var embalajes = resultados.get(mes).embalajes;
    var equipamientos = resultados.get(mes).equipamientos;
    var fijos = resultados.get(mes).fijos;
    var fletes = resultados.get(mes).fletes;
    var insumos = resultados.get(mes).insumos;
    var publicidad = resultados.get(mes).publicidad;
    var servicios = resultados.get(mes).servicios;
    var utilidadAntesImpuestos = resultados.get(mes).utilidadAntesImpuestos;
    var impuestos = resultados.get(mes).impuestos;
    var resultado = resultados.get(mes).resultado;
    var margenDeGanancia = resultados.get(mes).margenDeGanancia; // Obtener el margen de ganancia
    resultadosSheet.getRange("A" + row).setValue(mes);
    resultadosSheet.getRange("B" + row).setValue(ingresos);
    resultadosSheet.getRange("C" + row).setValue(egresos);
    resultadosSheet.getRange("D" + row).setValue(materiasPrimas);
    resultadosSheet.getRange("E" + row).setValue(manoDeObra);
    resultadosSheet.getRange("F" + row).setValue(comisiones);
    resultadosSheet.getRange("G" + row).setValue(envios);
    resultadosSheet.getRange("H" + row).setValue(utilidadBruta);
    resultadosSheet.getRange("I" + row).setValue(costosFinancieros);
    resultadosSheet.getRange("J" + row).setValue(embalajes);
    resultadosSheet.getRange("K" + row).setValue(equipamientos);
    resultadosSheet.getRange("L" + row).setValue(fijos);
    resultadosSheet.getRange("M" + row).setValue(fletes);
    resultadosSheet.getRange("N" + row).setValue(insumos);
    resultadosSheet.getRange("O" + row).setValue(publicidad);
    resultadosSheet.getRange("P" + row).setValue(servicios);
    resultadosSheet.getRange("Q" + row).setValue(utilidadAntesImpuestos);
    resultadosSheet.getRange("R" + row).setValue(impuestos);
    resultadosSheet.getRange("S" + row).setValue(resultado); 
    resultadosSheet.getRange("T" + row).setValue(margenDeGanancia); // Escribir el margen de ganancia en la columna correspondiente
    row++;
  });
}
