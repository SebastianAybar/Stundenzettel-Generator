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

    // Haupt-Read-Methode: Hier wird die Excel ausgelesen, die Monate geordnet, nach Monaten in eine große Liste von
    // einzelnen Listen aufgeteilt und die einzelnen Listen nach Nachname sortiert
    public List<List<MitarbeiterMonat>> getListsOfAbrechnungsmonate() {
        readAllRows();
        orderByAbrechnungsmonat(listMitarbeiterMonat);
        printList(listMitarbeiterMonat);
        System.out.println();
        System.out.println();

        List<List<MitarbeiterMonat>> lists = new ArrayList<>();

        List<MitarbeiterMonat> abrechnungsmonat;
        while((abrechnungsmonat = getNextAbrechnungsmonat()) != null) {
            orderByNachname(abrechnungsmonat);
            lists.add(abrechnungsmonat);
        }
        return lists;
    }

    // Methode, um den Excel-Inhalt in eine Liste von Monaten zu speichern
    public void readAllRows() {
        try (FileInputStream inputStream = new FileInputStream(inputPath)) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            listMitarbeiterMonat = new ArrayList<>();

            // Für jede Row wird ein Objekt erstellt, gefüllt und in die Liste aller Monate gespeichert
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

            // Richtige Werte beginnen in der Excel erst ab der zweiten Reihe, deswegen wird die Erste gelöscht
            listMitarbeiterMonat.remove(0);

            workbook.close();
        } catch (IOException e) {
            System.out.println("Error. Check code!");
        }
    }

    // Methode, um aus einer Cell den String-Wert zu bekommen und null-Exceptions zu verhindern durch "---" return
    private static String getCellAsString(Cell cell) {
        if (cell != null) return cell.toString();
        else return "---";
    }

    // Methode, um eine Liste von Monaten in der Konsole auszugeben
    public void printList(List<MitarbeiterMonat> list) {
        for (MitarbeiterMonat mitarbeiterMonat : list) {
            System.out.println(mitarbeiterMonat.toString());
        }
    }

    // Methode, um eine Liste von Monaten nach Abrechnungsdatum zu ordnen
    public void orderByAbrechnungsmonat(List<MitarbeiterMonat> list) {
        Comparator<MitarbeiterMonat> comparator = Comparator.comparing(MitarbeiterMonat::getAbrechnungsmonat);
        list.sort(comparator);
    }

    // Methode, um eine Liste von Monaten nach Nachnamen zu ordnen
    public void orderByNachname(List<MitarbeiterMonat> list) {
        Comparator<MitarbeiterMonat> comparator = Comparator.comparing(MitarbeiterMonat::getNachnameVorname);
        list.sort(comparator);
    }

    // Methode, um eine Liste aller Mitarbeiter aus demselben Abrechnungsmonat zurückzugeben
    // (Monat, der beim Ersten Objekt steht)
    public List<MitarbeiterMonat> getNextAbrechnungsmonat() {
        try {
            List<MitarbeiterMonat> listAbrechnungsmonat = new ArrayList<>();
            String abrechnungsmonat = listMitarbeiterMonat.get(0).getAbrechnungsmonat();
            for (int i = 0; i < listMitarbeiterMonat.size(); i++) {
                System.out.println(abrechnungsmonat + " == " + listMitarbeiterMonat.get(i).getAbrechnungsmonat());
                if (listMitarbeiterMonat.get(i).getAbrechnungsmonat().equals(abrechnungsmonat)) {
                    listAbrechnungsmonat.add(listMitarbeiterMonat.removeFirst());
                    i--;
                } else break;
            }
            return listAbrechnungsmonat;
        } catch (IndexOutOfBoundsException exception) {
            System.out.println("Error! No more Abrechnungsmonate!");
            return null;
        }
    }



}
