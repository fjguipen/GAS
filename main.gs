function onOpen(){
var ui = SpreadsheetApp.getUi();
  //Creamos el menu con submenu
  ui.createMenu('Más herramientas')
      .addItem('Eliminar duplicados', 'eliminarDesdeMismaTabla')
      .addItem('Eliminar duplicados desde tabla', 'eliminarDesdeDistintaTabla')
      .addToUi();
}

function eliminarDesdeMismaTabla (){
  //Hoja sobre la que se ha de trabajar
  var sheet =  SpreadsheetApp.getActive().getActiveSheet();
  //Numero de filas a saltar (encabezados)
  var titleRows =  Browser.inputBox("Por favor, introduce numero de filas a saltar(encabezados):")
  var col = Browser.inputBox("Por favor, introduce numero de la columna de trabajo:")
  
  
  rmDuplicadosMismaHoja2(sheet,Number(titleRows), Number(col));
}

 function eliminarDesdeDistintaTabla () {
   var hojaA = SpreadsheetApp.getActive().getActiveSheet();
   var titleRows = Browser.inputBox("Por favor, introduce numero de filas a saltar(encabezados):");
   var hojaACol = Browser.inputBox("Por favor, indica la columna donde se encuentran los datos de trabajo:");
   var hojaB = Browser.inputBox("Por favor, introduce el nombre de la hoja con los datos de referencia:");
   hojaB = SpreadsheetApp.getActive().getSheetByName(hojaB);
   if (!hojaB){ // Validación de la hojaB, si no existe, abortamos
      Browser.msgBox("La hoja especificada no existe");
      
   } else {
     var hojaBCol = Browser.inputBox("Por favor, indica la columna donde se encuentran los datos de referencia:");
     
     rmDuplicadosDistintaHoja(hojaB, Number(hojaBCol), Number(hojaACol), Number(titleRows), hojaA);
   }  
 }


