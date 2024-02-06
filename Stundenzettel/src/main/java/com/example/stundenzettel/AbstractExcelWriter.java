package com.example.stundenzettel;

import de.focus_shift.jollyday.core.Holiday;
import de.focus_shift.jollyday.core.HolidayManager;
import de.focus_shift.jollyday.core.ManagerParameters;
import org.apache.poi.ss.formula.functions.Na;
import org.apache.poi.ss.usermodel.*;
import org.apache.commons.math3.distribution.NormalDistribution;

import javax.xml.transform.Source;
import java.io.*;
import java.security.spec.RSAOtherPrimeInfo;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.temporal.TemporalAdjuster;
import java.time.temporal.TemporalAdjusters;
import java.util.*;
import java.time.format.DateTimeFormatter;

import static java.util.Locale.GERMANY;


public class AbstractExcelWriter {

    private final String BUNDESLAND = "he";
    private final String pathTemplate = "C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Stundenzettel\\Stundenzettel_Vorlage.xlsx";
    //private final String pathTemplate = "E:\\zAndere\\GitRepos\\Stundenzettel-Generator\\Stundenzettel\\Stundenzettel_Vorlage.xlsx";
    private final String outputPath;

    AbstractExcelWriter(String outputPath) {
        this.outputPath = outputPath;
    }

    public void writeToExcel(List<List<MitarbeiterMonat>> jahresliste) {
        int counter = 1;
        for (List<MitarbeiterMonat> monatsliste : jahresliste) {
            try {
                InputStream inputStream = new FileInputStream(pathTemplate);
                Workbook workbook = WorkbookFactory.create(inputStream);
                Sheet currentSheet = workbook.getSheetAt(0);
                for (int i = 0; i < monatsliste.size(); i++) {
                    workbook.cloneSheet(0);
                    currentSheet = workbook.getSheetAt(i+1);
                    String[] datum = monatsliste.get(i).getAbrechnungsmonat().split("/");
                    List<LocalDate> datenDesMonats = getDatenDesMonats(datum);
                    List<Cell> arbeitszeitenCells = new ArrayList<>();
                    int counterTage = 0;
                    List<Row> rowsToRemove = new ArrayList<>();
                    for (Row row : currentSheet) {
                        //String aktuellerMontag = "";
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
                                    } else {
//                                        Row ueberfluessigerTag = cell.getRow();
//                                        int rowNum = ueberfluessigerTag.getRowNum();
//                                        System.out.println(rowNum);
//                                        workbook.getSheetAt(i + 1).removeRow(ueberfluessigerTag);
                                        rowsToRemove.add(cell.getRow());
                                        break;
                                    }
                                }
                                /*if (cellValue.equals("Montag")) {
                                    aktuellerMontag = cell.getRow().getCell(2).getStringCellValue();
                                }*/
                                if (cellValue.startsWith("<<Std")) {
                                    arbeitszeitenCells.add(cell);
                                }
                            }
                        }

//                        if (rowRemoved) {
//                            workbook.getSheetAt(i + 1).shiftRows(removedRowNum + 1, workbook.getSheetAt(i + 1).getLastRowNum(), -1);
//
//                        }

                        //System.out.println(aktuellerMontag);
                    }//System.out.println("________________________________________________");

                    for (Row row : rowsToRemove) {
                        currentSheet.removeRow(row);
                    }

                    //Wir erstellen ein Array mit den normalverteilten Arbeitszeiten
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

                    for (int j = 0; j < arrayArbeitstage.length; j++) {
                        arrayArbeitstage[j] = arrayArbeitstage[j] * (stundensatz / sum);
                    }

                    double sumAfter = 0;
                    for (double value : arrayArbeitstage) {
                        sumAfter += value;
                    }

                    //Wir erstellen ein array welches wir dann in die List<Cell> schreiben
                    String[] arrArbeitszeitenCells = new String[arbeitszeitenCells.size()];

                    Random randomNumberGen = new Random();
                    int randomNumber;
                    DecimalFormat decimalFormat = new DecimalFormat("###.##");

                    for (int j = 0; j < arrayArbeitstage.length; j++) {
                        randomNumber = randomNumberGen.nextInt(arbeitszeitenCells.size() - 1);
                        while (arrArbeitszeitenCells[randomNumber] != null) {
                            randomNumber = randomNumberGen.nextInt(arbeitszeitenCells.size() - 1);
                        }
                        arrArbeitszeitenCells[randomNumber] = decimalFormat.format(arrayArbeitstage[j]);
                    }

                    //Befüllen des Sheets
                    String hourMinutes;
                    int insgMinuten, stunden, minuten;

                    for (int k = 0; k < arbeitszeitenCells.size(); k++) {
                        hourMinutes = "";
                        if (arrArbeitszeitenCells[k] != null) {
                            arbeitszeitenCells.get(k).setCellValue(arrArbeitszeitenCells[k]);
                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 1).setCellValue(arrArbeitszeitenCells[k]);

                            LocalDate aufgezeichnetAm = arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 2).getLocalDateTimeCellValue().toLocalDate().with(TemporalAdjusters.next(DayOfWeek.MONDAY));

                            if (isDatumEinFeiertag(aufgezeichnetAm, Integer.parseInt(datum[0]))) {
                                aufgezeichnetAm = aufgezeichnetAm.with(TemporalAdjusters.next(DayOfWeek.MONDAY));
                            }
                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 2).setCellValue(aufgezeichnetAm);

                            String temp = arrArbeitszeitenCells[k].replace(",", ".");
                            insgMinuten = (int) (Double.parseDouble(temp) * 60);
                            stunden = (int) Double.parseDouble(temp);
                            minuten = insgMinuten % 60;

                            hourMinutes = hourMinutes + "0" + stunden + ":";
                            if (minuten >= 10) hourMinutes += minuten;
                            else hourMinutes += "0" + minuten;

                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 1).setCellValue(hourMinutes);
                        } else {
                            arbeitszeitenCells.get(k).setCellValue("");
                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 1).setCellValue("");
                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 2).setCellValue("");
                        }
                    }

                    //Excel Formeln werden nach dem Füllen der Felder noch einmal ausgeführt (z.B. für KW)
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    evaluator.evaluateAll();

                }
                try (FileOutputStream fileOutputStream = new FileOutputStream(outputPath + "\\test" + counter++ + ".xlsx")) {
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

        double randomValue;
        for (int i = 0; i < numArbeitstage; i++) {
            do {
                randomValue = normalDistribution.sample();
            } while (randomValue < 0.25 || randomValue > 4);
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

        Cell arbeitszeitCell = cell.getRow().getCell(cell.getColumnIndex() + 1);
        arbeitszeitCell.setCellValue("");

        Cell dezimalCell = cell.getRow().getCell(cell.getColumnIndex() + 2);
//        dezimalCell.removeFormula();
        dezimalCell.setCellValue("");

        Cell arbeitszeitNettoCell = cell.getRow().getCell(cell.getColumnIndex() + 3);
        arbeitszeitNettoCell.removeFormula();
        arbeitszeitNettoCell.setCellValue("");

        Cell aufgezeichnetAmCell = cell.getRow().getCell(cell.getColumnIndex() + 4);
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
