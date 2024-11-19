<?php
if (!class_exists('COM')) {
    die('COM no está habilitado en PHP');
}

try {
    $excel = new COM("Excel.Application");
    $excel->Visible = true;
    $workbook = $excel->Workbooks->Add();
    $worksheet = $workbook->Worksheets(1);
    $worksheet->Name = "Lista de Estudiantes";

    $worksheet->Cells(1, 1)->Value = "Nombre";
    $worksheet->Cells(1, 2)->Value = "Numero de Control";
    $worksheet->Cells(1, 3)->Value = "Semestre";

    $estudiantes = [
        ["Apolinar", "211230017", "7mo Semestre"],
        ["Lizbet", "211230018", "7mo Semestre"],
        ["Oscar", "211230019", "7mo Semestre"],
        ["Fernando", "201230011", "9no Semestre"],
        ["Manuel Alejandro", "211230010", "9no Semestre"],
    ];

    $row = 2;
    foreach ($estudiantes as $estudiante) {
        $worksheet->Cells($row, 1)->Value = $estudiante[0];
        $worksheet->Cells($row, 2)->Value = $estudiante[1];
        $worksheet->Cells($row, 3)->Value = $estudiante[2];
        $row++;
    }

    $worksheet->Range("A1:C1")->Font->Bold = true;
    $worksheet->Range("A1:C1")->Interior->Color = 0xFFFF00;

    $worksheet->Columns("A:C")->AutoFit();

    $filePath = "C:\\xampp\\htdocs\\ejercicio\\lista_estudiantes.xlsx";

    if (!file_exists(dirname($filePath))) {
        mkdir(dirname($filePath), 0777, true);
    }

    $workbook->SaveAs($filePath);

    $workbook->Close(false);
    $excel->Quit();

    echo "Archivo Excel guardado correctamente en: " . $filePath;

} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>
