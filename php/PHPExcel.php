<?php

require_once 'PHPExcel.php';

// Leer el archivo de Excel
$excel = PHPExcel_IOFactory::load('php/tickets.xlsx');
$worksheet = $excel->getActiveSheet();

// Obtener el último número de ticket
$last_row = $worksheet->getHighestRow();
$last_ticket_num = $worksheet->getCell('A'.$last_row)->getValue();

// Generar un nuevo número de ticket
$new_ticket_num = $last_ticket_num + 1;

// Guardar el nuevo ticket en el archivo de Excel
$date = date('Y-m-d');
$time = date('H:i:s');
$details = $_POST['ticket-details'];
$worksheet->setCellValue('A'.($last_row+1), $new_ticket_num);
$worksheet->setCellValue('B'.($last_row+1), $date);
$worksheet->setCellValue('C'.($last_row+1), $time);
$worksheet->setCellValue('D'.($last_row+1), $details);
$excel_writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$excel_writer->save('php/tickets.xlsx');

// Redirigir al usuario a la página de seguimiento del ticket
header('Location: seguimiento-ticket.php?ticket-num='.$new_ticket_num);
exit();

?>
