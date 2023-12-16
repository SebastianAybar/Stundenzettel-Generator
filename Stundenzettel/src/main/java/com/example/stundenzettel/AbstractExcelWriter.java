package com.example.stundenzettel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AbstractExcelWriter {
    private final String pathTemplate = "E:\\zAndere\\GitRepos\\Stundenzettel-Generator\\Stundenzettel\\src\\main\\resources\\com\\example\\stundenzettel\\Stundenzettel_Vorlage.xlsx";
//    private final String pathOutput;

    public AbstractExcelWriter(/*String pathOutput*/) {
//        this.pathOutput = pathOutput;
    }

    public void writeToExcel(List<MitarbeiterMonat> listOfAbrechnungsmonat) {
        for (MitarbeiterMonat monat : listOfAbrechnungsmonat) {

            try (InputStream inputStream = new FileInputStream(pathTemplate)) {
                Workbook workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0);

                Map<String, String> mapToKeywords = new HashMap<>();
                mapToKeywords.put("<<Mitarbeiter>>", monat.getNachnameVorname());
                mapToKeywords.put("<<Mitarbeiternummer>>", monat.getMitarbeiternummer());
                mapToKeywords.put("<<Abrechnungsmonat>>", monat.getAbrechnungsmonat());

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue().trim();

                            if (mapToKeywords.containsKey(cellValue)) {
                                cell.setCellValue(mapToKeywords.get(cellValue));
                            }
                        }
                    }
                }

                try (FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\MM\\Downloads\\test\\output.xlsx")) {
                    workbook.write(fileOutputStream);
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("ERROR! LOOK AT CODE!");
            }
        break;
        }
    }

    public void convertToPdf() {

    }
}
