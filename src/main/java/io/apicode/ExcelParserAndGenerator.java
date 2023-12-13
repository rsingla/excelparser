package io.apicode;

import org.apache.poi.ss.usermodel.*;
        import org.apache.poi.xssf.usermodel.XSSFWorkbook;
        import org.apache.poi.ss.usermodel.Cell;
        import org.apache.poi.ss.usermodel.Row;

        import java.io.FileInputStream;
        import java.io.FileOutputStream;
        import java.io.IOException;

public class ExcelParserAndGenerator {
    public static void main(String[] args) {
        try {
            // Step 1: Read the Excel file
            FileInputStream inputFile = new FileInputStream("input.xlsx");
            Workbook workbook = new XSSFWorkbook(inputFile);
            Sheet sheet = workbook.getSheetAt(0); // assuming you want the first sheet

            // Step 2: Process the data
            // Example: Let's copy the content of the input sheet to a new sheet
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("outputSheet");

            int rowNum = 0;
            for (Row row : sheet) {
                Row newRow = newSheet.createRow(rowNum++);
                int cellNum = 0;
                for (Cell cell : row) {
                    Cell newCell = newRow.createCell(cellNum++);
                    switch (cell.getCellType()) {
                        case STRING:
                            newCell.setCellValue(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            newCell.setCellValue(cell.getNumericCellValue());
                            break;
                        // Add more cell type handling as needed
                    }
                }
            }

            // Step 3: Write to a new Excel file
            FileOutputStream outputFile = new FileOutputStream("output.xlsx");
            newWorkbook.write(outputFile);
            outputFile.close();

            // Step 4: Close input and output streams
            inputFile.close();

            System.out.println("Excel file processing completed.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}