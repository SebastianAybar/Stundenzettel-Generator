package com.example.stundenzettel;

import de.focus_shift.jollyday.core.Holiday;
import de.focus_shift.jollyday.core.HolidayManager;
import de.focus_shift.jollyday.core.ManagerParameters;
import org.apache.poi.ss.usermodel.*;
import org.apache.commons.math3.distribution.NormalDistribution;

import javax.xml.transform.Source;
import java.io.*;
import java.security.spec.RSAOtherPrimeInfo;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.*;
import java.time.format.DateTimeFormatter;

import static java.util.Locale.GERMANY;


public class AbstractExcelWriter {

    private final String BUNDESLAND = "he";
    private final String pathTemplate = "C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Stundenzettel\\Stundenzettel_Vorlage.xlsx";
    private final String outputPath;

    AbstractExcelWriter(String outputPath) {
        this.outputPath = outputPath;
    }

    public void writeToExcel(List<List<MitarbeiterMonat>> jahresliste) {
        for (List<MitarbeiterMonat> monatsliste : jahresliste) {
            try {
                InputStream inputStream = new FileInputStream(pathTemplate);
                Workbook workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0);
                for (int i = 0; i < monatsliste.size(); i++) {
                    workbook.cloneSheet(0);
                    String[] datum = monatsliste.get(i).getAbrechnungsmonat().split("/");
                    List<LocalDate> datenDesMonats = getDatenDesMonats(datum);
                    List<Cell> arbeitszeitenCells = new ArrayList<>();
                    //System.out.println(datenDesMonats);
                    int counterTage = 0;
                    for (Row row : workbook.getSheetAt(i + 1)) {
                        for (Cell cell : row) {
                            if (cell.getCellType() == CellType.STRING) {
                                String cellValue = cell.getStringCellValue().trim();
                                if (cellValue.equals("<<Abrechnungsmonat>>")) {
                                    cell.setCellValue(monatsliste.get(i).getAbrechnungsmonat());
                                }
                                if (cellValue.equals("<<Mitarbeiter>>")) {
                                    cell.setCellValue(monatsliste.get(i).getNachnameVorname());
                                }
                                if (cellValue.equals("<<Mitarbeiternummer>>")) {
                                    cell.setCellValue(monatsliste.get(i).getMitarbeiternummer());
                                }

                                if (cellValue.startsWith("<<Tag")) {
                                    if (counterTage < datenDesMonats.size()) {
                                        cell.setCellValue(datenDesMonats.get(counterTage));

                                        DateTimeFormatter deutschFormatierer = DateTimeFormatter.ofPattern("EEEE", Locale.GERMAN);
                                        String wochentag = datenDesMonats.get(counterTage).format(deutschFormatierer);
                                        Cell wochentagCell = cell.getRow().getCell(cell.getColumnIndex() - 1);
                                        wochentagCell.setCellValue(wochentag);

                                        if (wochentag.equals("Sonntag") || isDatumEinFeiertag(datenDesMonats.get(counterTage), Integer.parseInt(datum[0]))) {
                                            markiereRowAlsFreienTag(workbook, row, cell);
                                        }
                                        counterTage++;
                                    }
                                }
                                if (cellValue.startsWith("<<Std")) {
                                    arbeitszeitenCells.add(cell);
                                }
                            }
                        }
                    }

                    // Stundensatz berechnen
                    double svBrutto = Double.parseDouble(monatsliste.get(i).getSvBrutto());
                    double stundenlohn = 12;
                    double stundensatz = svBrutto / stundenlohn;
                    double arbeitstage = stundensatz / 2.5;
                    int gerundeteArbeitstage = (int) Math.round(arbeitstage);

                    double[] arrayArbeitstage = generateRandomNumbers(gerundeteArbeitstage, 2.5, 1);

                    double sum = 0;
                    for (double value : arrayArbeitstage) {
                        sum += value;
                    }
                    //System.out.println(sum);

                    for (int j = 0; j < arrayArbeitstage.length; j++) {
                        arrayArbeitstage[j] = arrayArbeitstage[j] * (stundensatz / sum);
                    }

                    double sumAfter = 0;
                    for (double value : arrayArbeitstage) {
                        sumAfter += value;
                    }
                    System.out.println(sumAfter);

                    for (double value : arrayArbeitstage) {
                        System.out.println(value);
                    }
                    System.out.println("_______________________________________________________________________________________________________________");

                    /*
                    int anzahlPotentielleArbeitstage = arbeitszeitenCells.size();
                    double[] arrArbeitszeitenCells = new double[anzahlPotentielleArbeitstage];
                    int[] zufallsIndizes = generateRandomIndices(anzahlPotentielleArbeitstage, gerundeteArbeitstage);
                    for (int k : zufallsIndizes) {
                        arrArbeitszeitenCells[k] = arrayArbeitstage[k];
                    }
                    for (double value : arrArbeitszeitenCells) {
                        System.out.println(value);
                    }
                    */

                    // Excel Formeln werden nach dem Füllen der Felder noch einmal ausgeführt (z.B. für KW)
                    //FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    //evaluator.evaluateAll();

                }
                try (FileOutputStream fileOutputStream = new FileOutputStream(outputPath + "\\test.xlsx")) {
                    workbook.write(fileOutputStream);
                }
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    private double[] generateRandomNumbers(int numArbeitstage, double mean, double sd) {
        double[] result = new double[numArbeitstage];
        NormalDistribution normalDistribution = new NormalDistribution(mean, sd);

        for (int i = 0; i < numArbeitstage; i++) {
            double randomValue = Math.max(normalDistribution.sample(), 0);
            result[i] = randomValue;
        }
        return result;
    }

    private static int[] generateRandomIndices(int arraySize, int numIndices) {
        int[] indices = new int[numIndices];
        Random random = new Random();

        for (int i = 0; i < numIndices; i++) {
            indices[i] = random.nextInt(arraySize);
        }
        return indices;
    }



    // Gibt eine Liste mit allen Tagen des bestimmten Monats zurück
    private List<LocalDate> getDatenDesMonats(String[] datum) {
        YearMonth jahrMonat = YearMonth.of(Integer.parseInt(datum[0]), Integer.parseInt(datum[1]));
        int anzahlTageImMonat = jahrMonat.lengthOfMonth();
        LocalDate ersterTag = jahrMonat.atDay(1);
        List<LocalDate> datenDesMonats = new ArrayList<>();
        for (int j = 0; j < anzahlTageImMonat; j++) {
            LocalDate aktuellesDatum = ersterTag.plusDays(j);
            datenDesMonats.add(aktuellesDatum);
        }
        return datenDesMonats;
    }

    // Row als freien Tag (Sonn- und Feiertage) markieren (-2, -1, +2, +3, +4, +5 um KW, Wochentag, etc. anzusprechen)
    private void markiereRowAlsFreienTag(Workbook workbook, Row row, Cell cell) {
        // Zellen leeren
        Cell kwCell = cell.getRow().getCell(cell.getColumnIndex() - 2);

        Cell wochentagCell = cell.getRow().getCell(cell.getColumnIndex() - 1);

        Cell arbeitszeitCell = cell.getRow().getCell(cell.getColumnIndex() + 2);
        arbeitszeitCell.setCellValue("");

        Cell dezimalCell = cell.getRow().getCell(cell.getColumnIndex() + 3);
        dezimalCell.removeFormula();
        dezimalCell.setCellValue("");

        Cell arbeitszeitNettoCell = cell.getRow().getCell(cell.getColumnIndex() + 4);
        arbeitszeitNettoCell.removeFormula();
        arbeitszeitNettoCell.setCellValue("");

        Cell aufgezeichnetAmCell = cell.getRow().getCell(cell.getColumnIndex() + 5);
        aufgezeichnetAmCell.setCellValue("");

        // Zellen färben
        CellStyle originalStyle = cell.getCellStyle();
        CellStyle freierTagStyle = workbook.createCellStyle();
        freierTagStyle.cloneStyleFrom(originalStyle);

        freierTagStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        freierTagStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        freierTagStyle.setBorderBottom(BorderStyle.THIN);
        freierTagStyle.setBorderLeft(BorderStyle.THIN);
        freierTagStyle.setBorderRight(BorderStyle.THIN);
        freierTagStyle.setBorderTop(BorderStyle.THIN);

        // Feld: KW
        CellStyle kwStyle = workbook.createCellStyle();
        kwStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        kwStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        kwStyle.setBorderBottom(BorderStyle.THIN);
        kwStyle.setBorderLeft(BorderStyle.THIN);
        kwStyle.setBorderRight(BorderStyle.THIN);
        kwStyle.setBorderTop(BorderStyle.THIN);
        Font dickFont = workbook.createFont();
        dickFont.setBold(true);
        kwStyle.setFont(dickFont);
        kwCell.setCellStyle(kwStyle);

        // Feld: Wochentag
        wochentagCell.setCellStyle(freierTagStyle);

        // Feld: Datum
        cell.setCellStyle(freierTagStyle);

        // Feld: Arbeitszeit
        arbeitszeitCell.setCellStyle(freierTagStyle);

        // Feld: Dezimal
        dezimalCell.setCellStyle(freierTagStyle);

        // Feld: Arbeitszeit netto
        arbeitszeitNettoCell.setCellStyle(freierTagStyle);

        // Feld: Aufgezeichnet am
        aufgezeichnetAmCell.setCellStyle(freierTagStyle);
    }

    // Prüft, ob das angegebene Datum ein Feiertag ist (mithilfe der Library <jollyday>
    private boolean isDatumEinFeiertag(LocalDate datum, int jahr) {
        final HolidayManager feiertageManager = HolidayManager.getInstance(ManagerParameters.create(GERMANY));
        final Set<Holiday> feiertage = feiertageManager.getHolidays(jahr, BUNDESLAND);
        for (Holiday feiertag : feiertage) {
            if (feiertag.getDate().toString().equals(datum.toString())) {
                return true;
            }
        }
        return false;
    }
}
