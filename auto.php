<?php
$outlook = new ("Outlook.Application");

$mail = $outlook->CreateItem(0);

$mail->To = "211230019@smarcos.tecnm.mx";
$mail->CC = "211230017@smarcos.tecnm.mx";
$mail->Subject = "Prueba de automatizacion";
$mail->Body = "DOCUMENTO DE CLASE";

$attachment = $mail->Attachments->Add("C:\Users\OSCAR\Documents\Semestre Agos-Dic 2024\Fundamentos de TI\Unidad 03/U3. Informe.pdf");


$mail->Send();
$outlook->Quit();
?>