<?php
defined('_JEXEC') or die('Restricted access');

$sheet = $phpexcel->getActiveSheet();

// Setup header row
$char = $reset_col_char;

foreach ($headers as $header) {
	$sheet->getStyle($char . 1)->applyFromArray(array_merge($border_box, $font_bold, $align_center));
	$sheet->setCellValue($char++ . 1, $header);
}

$row_index = $reset_data_row_index;
foreach($contacts as $contact)
{	
	$char = $reset_col_char;

	$sheet->getStyle($char . $row_index)->applyFromArray($border_box);
    $sheet->setCellValue($char++ . $row_index, $contact->first_name);

	$sheet->getStyle($char . $row_index)->applyFromArray($border_box);
    $sheet->setCellValue($char++ . $row_index, ucfirst($contact->last_name));

	$sheet->getStyle($char . $row_index)->applyFromArray($border_box);
    $sheet->setCellValue($char++ . $row_index, $contact->email);

	$sheet->getColumnDimension($char)->setAutoSize(true);
    $row_index++;
}

// Calulate the column widths to attempt to autosize them
$sheet->calculateColumnWidths();