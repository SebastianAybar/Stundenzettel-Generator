package com.example.stundenzettel;

import de.focus_shift.jollyday.core.Holiday;
import de.focus_shift.jollyday.core.HolidayManager;
import de.focus_shift.jollyday.core.ManagerParameters;
import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.WeekFields;
import java.util.*;

import static java.util.Locale.GERMANY;

public class EinzelerstellungReader {

    private final String pathTemplate = "C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Stundenzettel\\Stundenzettel_Vorlage.xlsx";
    private final String BUNDESLAND = "he";

    private final String abrechnungsmonat;
    private final String mitarbeiternummer;
    private final String svBrutto;
    private final String name;

    EinzelerstellungReader(String abrechnungsmonat, String mitarbeiternummer, String svBrutto, String name) {
        this.abrechnungsmonat = abrechnungsmonat;
        this.mitarbeiternummer = mitarbeiternummer;
        this.svBrutto = svBrutto;
        this.name = name;
    }

    public void writeToExcelEinzelerstellung() {
        try {
            InputStream inputStream = new FileInputStream(pathTemplate);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet currentSheet = workbook.getSheetAt(0);

            List<Cell> arbeitszeitenCells = new ArrayList<>();
            String[] datum = abrechnungsmonat.split("/");
            List<LocalDate> datenDesMonats = getDatenDesMonats(datum);
            List<Row> rowsToRemove = new ArrayList<>();

            //Wir befüllen die Felder Abrechnungsmonat, Mitarbeiter, Mitarbeiternummer und die Spalten Datum, KW und Wochentag
            int counterTage = 0;
            for (Row row : currentSheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue().trim();
                        if (cellValue.equals("<<Abrechnungsmonat>>")) {
                            cell.setCellValue(abrechnungsmonat);
                        }
                        if (cellValue.equals("<<Mitarbeiter>>")) {
                            cell.setCellValue(name);
                        }
                        if (cellValue.equals("<<Mitarbeiternummer>>")) {
                            cell.setCellValue(mitarbeiternummer);
                        }
                        if (cellValue.startsWith("<<Tag")) {
                            if (counterTage < datenDesMonats.size()) {
                                //Wir befüllen die Spalte Datum
                                cell.setCellValue(datenDesMonats.get(counterTage));
                                //Wir befüllen die Spalte KW
                                cell.getRow().getCell(cell.getColumnIndex() - 2).setCellValue(datenDesMonats.get(counterTage).get(WeekFields.of(Locale.getDefault()).weekOfWeekBasedYear()));
                                //Wir befüllen die Spalte Wochentag
                                DateTimeFormatter deutschFormatierer = DateTimeFormatter.ofPattern("EEEE", Locale.GERMAN);
                                String wochentag = datenDesMonats.get(counterTage).format(deutschFormatierer);
                                Cell wochentagCell = cell.getRow().getCell(cell.getColumnIndex() - 1);
                                wochentagCell.setCellValue(wochentag);
                                //Wir markieren und leeren die Sonn- und Feiertage
                                if (wochentag.equals("Sonntag") || isDatumEinFeiertag(datenDesMonats.get(counterTage), Integer.parseInt(datum[0]))) {
                                    markiereRowAlsFreienTag(workbook, row, cell);
                                }
                                counterTage++;
                            } else {
                                rowsToRemove.add(cell.getRow());
                                break;
                            }
                        }
                        if (cellValue.startsWith("<<Std")) {
                            arbeitszeitenCells.add(cell);
                        }

                    }
                }
            }

            for (Row row : rowsToRemove) {
                currentSheet.removeRow(row);
            }

            double svBrutto = Double.parseDouble(this.svBrutto);
            double stundenlohn = 12;
            double stundensatz = svBrutto / stundenlohn;
            double arbeitstage = stundensatz / 2.5;
            int gerundeteArbeitstage = (int) Math.round(arbeitstage);

            //Wir erstellen ein Array mit den Arbeitszeiten
            double[] arbeitszeiten = generateRandomNumbers(gerundeteArbeitstage, 2.5, 1);
            double sum = 0;
            for (double value : arbeitszeiten) {
                sum += value;
            }
            for (int j = 0; j < arbeitszeiten.length; j++) {
                arbeitszeiten[j] = arbeitszeiten[j] * (stundensatz / sum);
            }
            double sumAfter = 0;
            for (double value : arbeitszeiten) {
                sumAfter += value;
            }

            String[] werktage = new String[arbeitszeitenCells.size()];
            Random randomNumberGen = new Random();
            int randomNumber;
            DecimalFormat decimalFormat = new DecimalFormat("###.##");

            for (int j = 0; j < arbeitszeiten.length; j++) {
                randomNumber = randomNumberGen.nextInt(arbeitszeitenCells.size() - 1);
                while (werktage[randomNumber] != null) {
                    randomNumber = randomNumberGen.nextInt(arbeitszeitenCells.size() - 1);
                }
                werktage[randomNumber] = decimalFormat.format(arbeitszeiten[j]);
            }

            //Wir befüllen die Spalten, Dezimal, Arbeitszeit Netto, Aufgezeichnet am, und Arbeitszeit
            String hourMinutes;
            int insgMinuten, stunden, minuten;
            for (int i = 0; i < arbeitszeitenCells.size(); i++) {
                hourMinutes = "";
                if (werktage[i] != null) {
                    //Wir befüllen die Spalte "Dezimal"
                    arbeitszeitenCells.get(i).setCellValue(werktage[i]);
                    //Wir befüllen sie Spalte "Arbeitszeit Netto"
                    arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 1).setCellValue(werktage[i]);
                    //Wir befüllen die Spalte "Aufgezeichnet am"
                    LocalDate aufgezeichnetAm = arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 2).getLocalDateTimeCellValue().toLocalDate().with(TemporalAdjusters.next(DayOfWeek.MONDAY));
                    if (isDatumEinFeiertag(aufgezeichnetAm, Integer.parseInt(datum[0]))) {
                        aufgezeichnetAm = aufgezeichnetAm.with(TemporalAdjusters.next(DayOfWeek.MONDAY));
                    }
                    arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 2).setCellValue(aufgezeichnetAm);
                    //Wir befüllen die Spalte "Arbeitszeit"
                    String temp = werktage[i].replace(",", ".");
                    insgMinuten = (int) (Double.parseDouble(temp) * 60);
                    stunden = (int) Double.parseDouble(temp);
                    minuten = insgMinuten % 60;

                    hourMinutes = hourMinutes + "0" + stunden + ":";
                    if (minuten >= 10) hourMinutes += minuten;
                    else hourMinutes += "0" + minuten;

                    arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 1).setCellValue(hourMinutes);

                } else {
                    arbeitszeitenCells.get(i).setCellValue("");
                    arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 1).setCellValue("");
                    arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 2).setCellValue("");
                }
            }

            try (FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Documents\\test.xlsx")) {
                workbook.write(fileOutputStream);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
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

    // Row als freien Tag (Sonn- und Feiertage) markieren (-2, -1, +2, +3, +4, +5 um KW, Wochentag, etc. anzusprechen)
    private void markiereRowAlsFreienTag(Workbook workbook, Row row, Cell cell) {
        // Zellen leeren
        Cell kwCell = cell.getRow().getCell(cell.getColumnIndex() - 2);

        Cell wochentagCell = cell.getRow().getCell(cell.getColumnIndex() - 1);

        Cell arbeitszeitCell = cell.getRow().getCell(cell.getColumnIndex() + 1);
        arbeitszeitCell.setCellValue("");

        Cell dezimalCell = cell.getRow().getCell(cell.getColumnIndex() + 2);
        dezimalCell.setCellValue("");

        Cell arbeitszeitNettoCell = cell.getRow().getCell(cell.getColumnIndex() + 3);
        arbeitszeitNettoCell.removeFormula();
        arbeitszeitNettoCell.setCellValue("");

        Cell aufgezeichnetAmCell = cell.getRow().getCell(cell.getColumnIndex() + 4);
        aufgezeichnetAmCell.setCellValue("");

        // Zellen färben
        CellStyle originalStyle = cell.getCellStyle();
        CellStyle freierTagStyle = workbook.createCellStyle();
        CellStyle freierTagStyleFuerDatum = workbook.createCellStyle();
        freierTagStyle.cloneStyleFrom(originalStyle);

        freierTagStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        freierTagStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        freierTagStyle.setBorderBottom(BorderStyle.THIN);
        freierTagStyle.setBorderLeft(BorderStyle.THIN);
        freierTagStyle.setBorderRight(BorderStyle.THIN);
        freierTagStyle.setBorderTop(BorderStyle.THIN);

        freierTagStyleFuerDatum.cloneStyleFrom(freierTagStyle);

        freierTagStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("General"));

        // Feld: KW
        kwCell.setCellStyle(freierTagStyle);

        // Feld: Wochentag
        wochentagCell.setCellStyle(freierTagStyle);

        // Feld: Datum
        cell.setCellStyle(freierTagStyleFuerDatum);

        // Feld: Arbeitszeit
        arbeitszeitCell.setCellStyle(freierTagStyle);

        // Feld: Dezimal
        dezimalCell.setCellStyle(freierTagStyle);

        // Feld: Arbeitszeit netto
        arbeitszeitNettoCell.setCellStyle(freierTagStyle);

        // Feld: Aufgezeichnet am
        aufgezeichnetAmCell.setCellStyle(freierTagStyle);
    }
}
