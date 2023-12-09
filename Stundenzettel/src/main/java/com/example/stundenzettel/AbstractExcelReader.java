package com.example.stundenzettel;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class AbstractExcelReader {

//    private static final String path = StundenzettelController.getInputPath();
    public static void main(String[] args) {
//      System.out.println(path);
//      Pfad zur Excel-Datei
        String excelFilePath = "C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Side_Project_IFI\\Documents\\Mini-Job geringfügig Beschäftigte_01_10_2023_LV_Testnamen.xlsx";

        try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath))) {
            // Öffne die Arbeitsmappe (Workbook)
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Wähle das erste Blatt aus
            Sheet sheet = workbook.getSheetAt(0);

            // Iteriere durch die Zeilen und Spalten, um auf die Zellen zuzugreifen
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Hier kannst du auf den Zellwert zugreifen
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        default:
                            System.out.print("Nicht unterstützter Zelltyp\t");
                    }
                }
                System.out.println(); // Neue Zeile für neue Excel-Zeile
            }

            // Schließe die Arbeitsmappe (Workbook)
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
