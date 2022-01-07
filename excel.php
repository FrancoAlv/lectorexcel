<?php

#require_once "vendor/autoload.php";
# Indicar que usaremos el IOFactory
#use PhpOffice\PhpSpreadsheet\IOFactory;

include ("leectodeexcel.php");

# Recomiendo poner la ruta absoluta si no está junto al script
# Nota: no necesariamente tiene que tener la extensión XLSX
#$rutaArchivo = "002\\1_StandardReport.xlsx";
#$documento = IOFactory::load($rutaArchivo);

# Recuerda que un documento puede tener múltiples hojas
# obtener conteo e iterar
#$totalDeHojas = $documento->getSheetCount();


# Iterar hoja por hoja
#for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {
# Obtener hoja en el índice que vaya del ciclo
#$hojaActual = $documento->getSheet($indiceHoja);

$excel=new leectodeexcel("002\\1_StandardReport.xlsx");
$excel->start();


/*
$hojaActual = $documento->getSheetByName("Reporte de Asistencia");

#echo "<h3>Vamos en la hoja con índice $hojaActual</h3>";

foreach ($hojaActual->getRowIterator() as $fila) {
    foreach ($fila->getCellIterator() as $celda) {
        // Aquí podemos obtener varias cosas interesantes
        #https://phpoffice.github.io/PhpSpreadsheet/master/PhpOffice/PhpSpreadsheet/Cell/Cell.html

        # El valor, así como está en el documento
        $valorRaw = $celda->getValue();
        if ($valorRaw == "Periodo:") {
        $banderaperiodo=true;
        }

        # Formateado por ejemplo como dinero o con decimales
        $valorFormateado = $celda->getFormattedValue();

        # Si es una fórmula y necesitamos su valor, llamamos a:
        $valorCalculado = $celda->getCalculatedValue();

        # Fila, que comienza en 1, luego 2 y así...
        $fila = $celda->getRow();
        # Columna, que es la A, B, C y así...
        $columna = $celda->getColumn();

        echo "En <strong>$columna$fila</strong> tenemos el valor <strong>$valorRaw</strong>. ";
        echo "Formateado es: <strong>$valorFormateado</strong>. ";
        echo "Calculado es: <strong>$valorCalculado</strong><br><br>";
    }

    

}

*/

