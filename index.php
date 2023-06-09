<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
?>
<!DOCTYPE html>
<html>
<head>
    <title>Cargar archivo Excel y actualizar MySQL</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <h1>Actualiza existencias productos</h1>

    <?php
    // Verificar si se ha enviado un archivo
    if(isset($_FILES['archivo_excel'])) {

        // Obtener el archivo subido
        $archivo = $_FILES['archivo_excel']['tmp_name'];

        // Cargar el archivo Excel
        $spreadsheet = IOFactory::load($archivo);

        // Obtener la hoja activa
        $hoja = $spreadsheet->getActiveSheet();

        // Obtener los nombres de las columnas de la primera fila
        $columnas = [];
        $encabezados = $hoja->getRowIterator(1)->current()->getCellIterator();
        foreach ($encabezados as $celda) {
            $columnas[] = $celda->getValue();
        }


        // Verificar si las columnas "id" y "existencias" están presentes en el archivo Excel
if (in_array('codigo', $columnas) && in_array('existencia', $columnas)) {
    // Conexión a la base de datos MySQL
    $host = '5.161.124.186';
    $puerto = 3306;
    $usuario = 'acupunturacom_80';
    $contraseña = 'Luis1523348...$$';
    $base_datos = 'acupunturacom_70';

    $conexion = new mysqli($host, $usuario, $contraseña, $base_datos, $puerto);

    // Verificar la conexión
    if ($conexion->connect_error) {
        die("Error de conexión a MySQL: " . $conexion->connect_error);
    }

    // Recorrer las filas del archivo Excel (empezando desde la fila 2, asumiendo que la primera fila son los encabezados)
    foreach ($hoja->getRowIterator(2) as $fila) {
        $datos = [];
        $celdas = $fila->getCellIterator();
        foreach ($celdas as $celda) {
            $datos[] = $celda->getValue();
        }

        // Verificar si la fila tiene los datos de las columnas requeridas
        if (count($datos) >= 2) {
            // Actualizar la base de datos con los datos del archivo Excel
            $id = $conexion->real_escape_string($datos[array_search('codigo', $columnas)]);
            $existencias = $conexion->real_escape_string($datos[array_search('existencia', $columnas)]);

            $sql = "UPDATE acmx_product SET available_for_order = '{$existencias}' WHERE reference = '{$id}'";
            $resultado = $conexion->query($sql);

            if (!$resultado) {
                echo "Error al actualizar los datos: " . $conexion->error;
            }
        }
    }

    // Cerrar la conexión a la base de datos
    $conexion->close();

    echo "Datos actualizados correctamente.";
} else {
    echo "El archivo Excel no contiene las columnas requeridas.";
}
    }
    ?>

    <form method="POST" enctype="multipart/form-data">
    <label for="archivo_excel" class="file-input">
            Seleccionar archivo
            <input type="file" name="archivo_excel" id="archivo_excel" style="display: none;">
        </label>
        <input type="submit" value="Cargar">
    </form>
</body>
</html>

