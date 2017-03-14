<?php

/**
 * @file
 * Contains module theme functions.
 */

/**
 * Csv content builder function.
 */
function theme_vbo_export_csv($variables) {
  // Sanitize data.
  foreach ($variables['header'] as $key => $item) {
    $variables['header'][$key] = strtr($item, array($variables['separator'] => ' '));
  }

  $content_replacements = array(
    '\r\n' => ' ',
    '\n\r' => ' ',
    '\r' => ' ',
    '\n' => ' ',
    '\t' => ' ',
    $variables['separator'] => ' '
  );
  foreach ($variables['rows'] as $row_key => $row) {
    foreach ($row as $cell_key => $cell) {
      $variables['rows'][$row_key][$cell_key] = strtr($cell, $content_replacements);
    }
  }

  // Generate output.
  $csv_rows = array();
  $csv_rows[] = implode(';', $variables['header']);
  foreach ($variables['rows'] as $row) {
    $csv_rows[] = implode(';', $row);
  }

  return implode(PHP_EOL, $csv_rows);
}

/**
 * Xlsx content builder function.
 */
function theme_vbo_export_xlsx($variables) {
  // Sanitize data.
  $header = array();
  foreach ($variables['header'] as $item) {
    $header[] = $item;
  }

  $rows = array();
  $row_index = 0;
  foreach ($variables['rows'] as $row) {
    $rows[$row_index] = array();
    foreach ($row as $cell) {
      $rows[$row_index][] = $cell;
    }
    $row_index++;
  }

  // Load PHPExcel library.
  $library = libraries_load('PHPExcel');

  // Create PHPExcel spreadsheet and add rows to it.
  $objPHPExcel = new PHPExcel();
  $objPHPExcel->removeSheetByIndex();
  $objPHPExcel->getProperties()
    ->setCreated('VBO Export')
    ->setTitle('VBO Export - ' . date('d-m-Y H:i', REQUEST_TIME));

  $worksheet = $objPHPExcel->createSheet();
  $worksheet->setTitle(t('Export'));

  // Set header.
  foreach ($header as $col_index => $label) {
    $worksheet->setCellValueExplicitByColumnAndRow($col_index, 1, $label);
  }

  // Set rows.
  foreach ($rows as $row_index => $row) {
    // Rows start from 1 and we need to account for header.
    $row_index += 2;
    foreach ($row as $col_index => $cell) {
      $worksheet->setCellValueExplicitByColumnAndRow($col_index, $row_index, $cell);
    }
  }

  // Add some additional styling to the worksheet.
  $objPHPExcel->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
  $last_column = $worksheet->getHighestColumn();
  $last_column_index = PHPExcel_Cell::columnIndexFromString($last_column);

  // Define the range of the first row.
  $first_row_range = 'A1:' . $last_column . '1';

  // Set first row in bold.
  $worksheet->getStyle($first_row_range)->getFont()->setBold(TRUE);

  // Activate an autofilter on the first row.
  $worksheet->setAutoFilter($first_row_range);

  // Set wrap text and top vertical alignment for the entire worksheet.
  $full_range = 'A1:' . $last_column . $worksheet->getHighestRow();
  $worksheet->getStyle($full_range)->getAlignment()
    ->setWrapText(TRUE)
    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

  // Adjust the column size for each column.
  PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
  for ($column = 0; $column <= $last_column_index; $column++) {
    $worksheet->getColumnDimensionByColumn($column)->setAutoSize(TRUE);
  }

  // Set a minimum and maximum width for columns.
  $min_column_width = 15;
  $max_column_width = 85;
  $worksheet->calculateColumnWidths();
  for ($column = 0; $column <= $last_column_index; $column++) {
    $width = $worksheet->getColumnDimensionByColumn($column)->getWidth();
    if ($width < $min_column_width) {
      $worksheet->getColumnDimensionByColumn($column)->setAutoSize(FALSE);
      $worksheet->getColumnDimensionByColumn($column)->setWidth($min_column_width);
    }
    elseif ($width > $max_column_width) {
      $worksheet->getColumnDimensionByColumn($column)->setAutoSize(FALSE);
      $worksheet->getColumnDimensionByColumn($column)->setWidth($max_column_width);
    }
  }

  $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

  // Catch the output of the spreadsheet.
  ob_start();
  $objWriter->save('php://output');
  $excelOutput = ob_get_clean();
  return $excelOutput;
}
