/********************
Funcion que admite tres parametros:
[sheet : Sheet] -> La hoja sobre la que queremos trabajar [Si lo omitimos, cojerá la hoja activa]
[jumpRows : Int] -> Si nuestra hoja contiene cabeceras, indicaremos cuantas lineas deseamos saltar. [Por omision su valor es 0]
[colDRef : Int] -> Indica el nº de la columna que tomaremos como referencia para validar las filas [Por defecto, columna 1 (valor 0)]

Objetivo: Eliminar aquellas FILAS que estén duplicadas, obetniendo como referencia el valor de una determinada columna.
*********************/

function rmDuplicadosMismaHoja2 ( sheet, jumpRows, colDRef ) {
  //Declaración de variables
    //Nueva hoja para los datos
  var newSheet = sheet.getParent().insertSheet(sheet.getName()+"_UPDATED")
    //Datos de trabajo
  var originalSheetData = sheet.getRange(1+jumpRows, 1, sheet.getMaxRows()-jumpRows, sheet.getMaxColumns()).getValues();
    //Preparamos la estructura que contendrá los datos sin duplicar
  var newSheetData = [];
  
  /*
  Ordena los valores de la columna que queremos validar de modo que los elementos repetidos 
  queden colocados de manera consecutiva.
  */
  originalSheetData.sort(function(a,b){
    //Ponemos todo a minusculas para evitar errores
    var nameA = a[0].toLowerCase();
    var nameB = b[0].toLowerCase();
    //Para orden alfabético:
    if (nameA < nameB){
      return -1
    }
    if (nameA === nameB){
      return 0
    }
    return 1
  })
  
  
  /*
  Creamos una tabla nueva con aquellos elementos que no sean precedidos 
  por un elemento de igual valor
  */
    //Primero añadimos a nuestra nueva tabla las filas pertenecientes a las cabeceras
  for (var row = 0; row < jumpRows; row++){    
    newSheetData.push(sheet.getRange(row+1, 1, 1, sheet.getMaxColumns()).getValues()[0])
  }
  Logger.log(originalSheetData)
    //Segundo añadimos la primera fila de datos, ya que ésta por si misma no puede estar duplicada
  newSheetData.push(originalSheetData[0])
    //Comprobamos las sucesivas filas en busca de duplicidad
  for (var row = 1; row < originalSheetData.length; row++){   
      if (originalSheetData[row][colDRef-1] != originalSheetData[row-1][colDRef-1])
      {
        newSheetData.push(originalSheetData[row])
      }
  }
  Logger.log(newSheetData)
  //Finalmente insertamos la tabla en la nueva Hoja
  newSheet.getRange(1, 1, newSheetData.length, newSheetData[0].length).setValues(newSheetData);
}
