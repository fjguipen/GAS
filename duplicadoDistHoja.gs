/********************
Funcion que admite tres parametros:
hojaB : Sheet -> Hoja en la cual encontramos los datos
[hojaA : Sheet] -> La hoja sobre la que queremos trabajar [Si lo omitimos, cojerá la hoja activa]
[colDRefA : Int] -> Indica el nº de la columna que tomaremos como referencia para validar las filas [Por defecto, columna 1 (valor 0)]

Objetivo: Eliminar aquellas FILAS que estén duplicadas, obetniendo como referencia el valor de una determinada columna.
*********************/

function rmDuplicadosDistintaHoja(hojaB, hojaBCol, hojaACol, jumpRowsA, hojaA){
  //Declaración de variables
    //Datos de trabajo
  var dataA = hojaA.getRange(1+jumpRowsA, 1, hojaA.getMaxRows()-jumpRowsA, hojaA.getMaxColumns()).getValues();
  
  var dataB = hojaB.getRange(1, 1, hojaB.getMaxRows(), 1).getValues().map(function (e){
    return e.toString()
  });
  
  /*
  Realizamos un filtrado de los datos de A, 
  de modo que sólo nos quedamos con aquellos que no existan en B.
  */
  newSheetData = dataA.filter(function (e){
    return dataB.indexOf(e[ hojaACol - 1 ]) >= 0 ? false : true
  })
    
  //Si todo ha ido bien, Insertamos los datos generados en una nueva hoja.
  if (newSheetData){
    hojaA.getParent().insertSheet(hojaA.getName()+"_UPDATED_FS").getRange(1, 1, newSheetData.length, newSheetData[0].length).setValues(newSheetData);
  }
} 
