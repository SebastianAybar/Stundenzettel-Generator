package com.example.stundenzettel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class AbstractExcelReader {
    private final String inputPath;
    List<MitarbeiterMonat> listMitarbeiterMonat;
    AbstractExcelReader(String inputPath) {
        this.inputPath = inputPath;
    }

    public List<List<MitarbeiterMonat>> getListOfAbrechnungsmonate() {
        readAllRows();
        orderByAbrechnungsmonat(listMitarbeiterMonat);
        //printList(listMitarbeiterMonat);

        List<List<MitarbeiterMonat>> jahresliste = new ArrayList<>();

        List<MitarbeiterMonat> monatsliste = new ArrayList<>();
        String monat = listMitarbeiterMonat.get(0).getAbrechnungsmonat();
        for (MitarbeiterMonat row : listMitarbeiterMonat) {
            if (row.getAbrechnungsmonat().equals(monat)) {
                monatsliste.add(row);
            } else {
                jahresliste.add(monatsliste);
                monatsliste = new ArrayList<>();
                monat = row.getAbrechnungsmonat();
                monatsliste.add(row);
            }
        }
        jahresliste.add(monatsliste);
        //printJahresliste(jahresliste);
        System.out.println();
        return jahresliste;
    }

    public void readAllRows() {
        try (FileInputStream inputStream = new FileInputStream(inputPath)) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            listMitarbeiterMonat = new ArrayList<>();

            for (Row row : sheet) {
                MitarbeiterMonat mitarbeiterMonat = new MitarbeiterMonat();
                mitarbeiterMonat.setNr(getCellAsString(row.getCell(0)));
                mitarbeiterMonat.setBerater(getCellAsString(row.getCell(1)));
                mitarbeiterMonat.setMandant(getCellAsString(row.getCell(2)));
                mitarbeiterMonat.setAbrechnungsmonat(getCellAsString(row.getCell(3)));
                mitarbeiterMonat.setMitarbeiternummer(getCellAsString(row.getCell(4)));
                mitarbeiterMonat.setNachnameVorname(getCellAsString(row.getCell(5)));
                mitarbeiterMonat.setPgr(getCellAsString(row.getCell(6)));
                mitarbeiterMonat.setStkl(getCellAsString(row.getCell(7)));
                mitarbeiterMonat.setAntyp(getCellAsString(row.getCell(8)));
                mitarbeiterMonat.setBschluesselKv(getCellAsString(row.getCell(9)));
                mitarbeiterMonat.setBschluesselRv(getCellAsString(row.getCell(10)));
                mitarbeiterMonat.setBschluesselAv(getCellAsString(row.getCell(11)));
                mitarbeiterMonat.setBschluesselPv(getCellAsString(row.getCell(12)));
                mitarbeiterMonat.setSvBrutto(getCellAsString(row.getCell(13)));
                mitarbeiterMonat.setNameBetrieb(getCellAsString(row.getCell(14)));
                mitarbeiterMonat.setStrassePostfach(getCellAsString(row.getCell(15)));
                mitarbeiterMonat.setPlz(getCellAsString(row.getCell(16)));
                mitarbeiterMonat.setOrt(getCellAsString(row.getCell(17)));
                mitarbeiterMonat.setSvTage(getCellAsString(row.getCell(18)));
                mitarbeiterMonat.setNatKennzeichen(getCellAsString(row.getCell(19)));
                mitarbeiterMonat.setMidijobregelung(getCellAsString(row.getCell(20)));
                listMitarbeiterMonat.add(mitarbeiterMonat);
            }

            listMitarbeiterMonat.remove(0);

            workbook.close();
        } catch (IOException e) {
            System.out.println("Error. Check code!");
        }
    }
    private static String getCellAsString(Cell cell) {
        if (cell != null) return cell.toString();
        else return "---";
    }

    public void orderByAbrechnungsmonat(List<MitarbeiterMonat> list) {
        Comparator<MitarbeiterMonat> comparator = Comparator.comparing(MitarbeiterMonat::getAbrechnungsmonat);
        list.sort(comparator);
    }

    public void printList(List<MitarbeiterMonat> list) {
         for (MitarbeiterMonat mitarbeiterMonat : list) {
             System.out.println(mitarbeiterMonat.toString());
         }
    }

    public void printJahresliste(List<List<MitarbeiterMonat>> lists) {
        for (List<MitarbeiterMonat> list : lists) {
            printList(list);
            System.out.println();
        }
    }
}
