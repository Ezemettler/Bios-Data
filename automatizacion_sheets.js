function OrdenFunciones() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ventasSheet = spreadsheet.getSheetByName("Ventas");
    var ventasData = ventasSheet.getDataRange().getValues(); // Obtener datos de ventas
    var egresosSheet = spreadsheet.getSheetByName("Egresos");
    var egresosData = egresosSheet.getDataRange().getValues(); // Obtener datos de egresos
    var resultadosSheet = spreadsheet.getSheetByName("Resultados");
    
    // Escribir los resultados en la hoja "Resultados" primero
    resultadosSheet.clear(); // Limpiar la hoja antes de escribir nuevos datos
    var ingresos = Resultados_Ingresos(ventasData); // Obtener los ingresos
    var sumaCostoVenta = calcularSumaCostoVenta(ventasData); // Obtener la suma de costos de venta
    var sumaEgresos = calcularSumaEgresos(egresosData); // Obtener la suma de egresos
    
    // Calcular el total de egresos sumando el costo de venta
    sumaEgresos += sumaCostoVenta;
    
    // Escribir los resultados en la hoja "Resultados"
    var resultados = [['Mes', 'Ingresos', 'Egresos']];
    for (var mesAnio in ingresos) {
      resultados.push([mesAnio, ingresos[mesAnio], sumaEgresos]);
    }
    resultadosSheet.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);
    
    // Realizar otros cálculos después de actualizar los resultados
    Ventas_TotalACobrar();
    Ventas_CostoVenta();
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
  
  function Resultados_Ingresos(ventasData) {
    var ingresos = {};
    for (var i = 1; i < ventasData.length; i++) { // Cambiado a empezar desde la fila 0
      var fecha = new Date(ventasData[i][0]);
      var mesAnio = fecha.getFullYear() + "-" + (fecha.getMonth() + 1);
      var total = Number(ventasData[i][21]);
      if (!isNaN(total)) {
        ingresos[mesAnio] = (ingresos[mesAnio] || 0) + total;
      }
    }
    return ingresos; // Devolver los ingresos calculados
  }
  
  
  function calcularSumaCostoVenta(ventasData) {
      var sumaCostoVenta = 0;
      for (var i = 1; i < ventasData.length; i++) {
          var costoVenta = parseFloat(ventasData[i][22]); // Suponiendo que la columna del costo de venta es la 23 (W)
          if (!isNaN(costoVenta)) {
              sumaCostoVenta += costoVenta;
          }
      }
      return sumaCostoVenta; // Devolver la suma de costos de venta calculada
  }
  
  function calcularSumaEgresos(egresosData) {
      var sumaEgresos = 0;
      for (var i = 1; i < egresosData.length; i++) {
          var valorTotal = parseFloat(egresosData[i][4]);
          if (!isNaN(valorTotal)) {
              sumaEgresos += valorTotal;
          }
      }
    
      return sumaEgresos; // Devolver la suma de egresos calculada
  }
  
  