function separarTroquel() {
  let spreadsheet = SpreadsheetApp.getActive();
  console.log('separar troquel iniciado y ejecutado')
  // con esto separamos el nombre del producto y el GTIN 
  spreadsheet.getRange('Productos!C:C').splitTextToColumns(']');
  spreadsheet.getRange('Productos!C:C').splitTextToColumns('[');
}

function formatoCeldas() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  let spreadsheet = SpreadsheetApp.getActive();
  let impar = 1
  let par = 2
  let i = 0

  spreadsheet.getSheetByName('Rotulos').activate();
  sheet.getRange('Rotulos!A:B').setHorizontalAlignment("center");
  while (spreadsheet.getRange('Productos!A5').getValue() / 2 + 5 > i) {
    spreadsheet.setRowHeight(impar, 45);
    sheet.getRange('A' + impar + ':C' + impar).setVerticalAlignment('bottom');
    spreadsheet.getRange('A' + impar + ':C' + impar).setTextStyle(SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontFamily('Times New Roman')
      .setFontSize(13)
      .build());

    spreadsheet.setRowHeight(par, 80);
    sheet.getRange('A' + par + ':C' + par).setVerticalAlignment('top');
    spreadsheet.getRange('A' + par + ':C' + par).setTextStyle(SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontFamily('Times New Roman')
      .setFontSize(35)
      .build());
    //acumuladores
    i = i + 1
    impar = impar + 2
    par = par + 2
    console.log(i, ' - impar:', impar, ' - par: ', par)
  }
}

function limpiarPlanilla() {
  let spreadsheet = SpreadsheetApp.getActive();

  //Borra los datos de la hoja Productos, limpiando los productos y su viejo sector y 
  spreadsheet.getRange('Productos!B:E').clearContent();
  spreadsheet.getRange('Productos!A6:A').clearContent();

  spreadsheet.getRange('Productos!B1').setValue('PEGAR AQUI Hoja Completa').setTextStyle(SpreadsheetApp.newTextStyle()
    .setFontSize(15)
    .setItalic(true)
    .build())

  // Borra el contenido en la hoja Rotulos y Solo Nombres, y tambien borra los bordes
  spreadsheet.getRange('Rotulos!A:C')
    .clearContent()
    .setBorder(false, false, false, false, false, false);

  spreadsheet.getRange('Solo Nombres!A:B')
    .clearContent()
    .setBorder(false, false, false, false, false, false);

  spreadsheet.getRange('Productos!A6').setValue('Pegar aquí Sector Nuevo');

  spreadsheet.getRange('Productos!C6').setValue('Pegar aquí Nombre de Producto');

  spreadsheet.getRange('Productos!B1').activate();
}

function soloNombres() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = SpreadsheetApp.getActiveSheet().activate();
  let filas, i = 1

  spreadsheet.getSheetByName('Productos').activate();
  separarTroquel();

  spreadsheet.getRange('Productos!C6').activate();
  filas = spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getNumRows();
  console.log(filas)

  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getActiveRange().copyTo(spreadsheet.getRange('Solo Nombres!A1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  while ((filas + 1) != i) {
    spreadsheet.getSheetByName('Solo Nombres').setRowHeight(i, 45);
    i = i + 1
  }

  sheet.getRange('Solo Nombres!A:B').setHorizontalAlignment("center").setVerticalAlignment('bottom');
  spreadsheet.getRange('Solo Nombres!A:B').setTextStyle(SpreadsheetApp.newTextStyle()
    .setFontSize(13)
    .setFontFamily('Times New Roman')
    .setBold(true)
    .build())

  spreadsheet.getRange('Solo Nombres!A1').activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getActiveRange().setBorder(true, true, true, true, true, true, 'gray', null)

  spreadsheet.getSheetByName('Solo Nombres').activate();
}

function rotulador() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sector = spreadsheet.getRange('Productos!A6:A').getValues();
  let nombre = spreadsheet.getRange('Productos!C6:C').getValues();
  let filaNombre = 1;
  let filaSector = 2;

  formatoCeldas();

  // Función para convertir un array multidimensional en un array plano y eliminar los datos vacios = ''
  function filtrarArray(array) {
    let arrayFlat = array.flat();
    return arrayFlat.filter(element => element !== "" && element !== null);
  }

  // Ejecutamos la funcion en sector y nombre para tener el array limpio
  let sectorFiltrado = Object.values(filtrarArray(sector));
  let nombreFiltrado = Object.values(filtrarArray(nombre));

  let k = 0; // Arreglo para los arrays
  for (i = 0; i < sectorFiltrado.length / 3; i++) {
    for (j = 0; j < 3; j++) {
      // 
      let columna = ['A', 'B', 'C']
      console.log(columna[j] + filaNombre);
      console.log(columna[j] + filaSector);
      // Añadimos los valores de los arrays en la hoja Rotulos
      spreadsheet.getRange('Rotulos!' + columna[j] + filaNombre).setValue(nombreFiltrado[k])
      spreadsheet.getRange('Rotulos!' + columna[j] + filaSector).setValue(sectorFiltrado[k])
      k += 1
    }
    // Ponemos bordes a cada rotulo
    spreadsheet.getRange('Rotulos!A' + filaNombre + ':C' + filaSector).setBorder(true, true, true, true, true, false);

    // Acumuladores
    filaNombre += 2;
    filaSector += 2;
  }
  Utilities.sleep(500)
  createPDF(ssId, sheet, pdfName);
}
