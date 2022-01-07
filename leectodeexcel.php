<?php
require_once "vendor/autoload.php";
# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

class leectodeexcel
{

    public function __construct($ruta)
    {
        $this->rutaArchivo = $ruta;
        $this->documento = IOFactory::load($this->rutaArchivo);
    }

    private $banderaperiodo = false;

    private $banderafechaactual = false;

    private $banderaid = false;

    private $banderanombre = false;

    private $banderadepartamento = false;

    private $banderahoras = false;

    private $rutaArchivo = null;

    private $documento = null;


    private function  vereficadodebanderas($valor)
    {
        $valoRaw = $valor->getValue();
        if ($this->banderahoras) {
            echo "estas son las horas $valoRaw \n";
            return;
        }
        if ($this->banderaperiodo) {
            echo "este es un periodo  $valoRaw \n";
            $this->banderaperiodo = false;
            return;
        }
        if ($this->banderafechaactual) {
            echo "este es un fecha actual $valoRaw \n";
            $this->banderafechaactual = false;
            return;
        }
        if ($this->banderadepartamento) {
            echo "este es el departamento $valoRaw \n";
            $this->banderadepartamento = false;
            $this->banderahoras = true;
            return;
        }

        if ($this->banderaid) {
            echo "este es un id $valoRaw \n";
            $this->banderaid = false;
            return;
        }
        if ($this->banderanombre) {
            echo "este es un nombre $valoRaw \n";
            $this->banderanombre = false;
            return;
        }
    }


    public function start()
    {
        $hojaActual = $this->documento->getSheetByName("Reporte de Asistencia");

        #echo "<h3>Vamos en la hoja con índice $hojaActual</h3>";

        foreach ($hojaActual->getRowIterator() as $fila) {
            foreach ($fila->getCellIterator() as $celda) {
                // Aquí podemos obtener varias cosas interesantes
                #https://phpoffice.github.io/PhpSpreadsheet/master/PhpOffice/PhpSpreadsheet/Cell/Cell.html


                # El valor, así como está en el documento
                $valorRaw = $celda->getValue();

                if ($valorRaw != "" || $valorRaw != null) {
                    $this->vereficadodebanderas($celda);
                } else {
                   # $this->banderahoras = false;
                }

                if ($valorRaw == "Periodo:") {
                    $this->banderaperiodo = true;
                }

                if ($valorRaw == "Fecha actual:") {
                    $this->banderafechaactual = true;
                }

                if ($valorRaw == "ID:") {
                    $this->banderaid = true;
                    $this->banderahoras = false;
                }

                if ($valorRaw == "Nombre:") {
                    $this->banderanombre = true;
                }

                if ($valorRaw == "Departamento:") {
                    $this->banderadepartamento = true;
                }


                # Formateado por ejemplo como dinero o con decimales
                $valorFormateado = $celda->getFormattedValue();

                # Si es una fórmula y necesitamos su valor, llamamos a:
                $valorCalculado = $celda->getCalculatedValue();

                # Fila, que comienza en 1, luego 2 y así...
                $fila = $celda->getRow();
                # Columna, que es la A, B, C y así...
                $columna = $celda->getColumn();

                /* echo "En <strong>$columna$fila</strong> tenemos el valor <strong>$valorRaw</strong>. ";
                    echo "Formateado es: <strong>$valorFormateado</strong>. ";
                    echo "Calculado es: <strong>$valorCalculado</strong><br><br>";
                */
            }
        }
    }
}
