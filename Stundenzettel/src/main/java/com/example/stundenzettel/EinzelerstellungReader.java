package com.example.stundenzettel;

import de.focus_shift.jollyday.core.Holiday;
import de.focus_shift.jollyday.core.HolidayManager;
import de.focus_shift.jollyday.core.ManagerParameters;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.WeekFields;
import java.util.*;

import static com.example.stundenzettel.Attribute.DOCUMENT_FILE_SUFFIX;
import static com.example.stundenzettel.Attribute.PATH_DATEI_STUNDENZETTELVORLAGE;
import static java.util.Locale.GERMANY;

public class EinzelerstellungReader {

    //    private final String pathTemplate = PATH_DATEI_STUNDENZETTELVORLAGE;
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

    public void writeToExcelEinzelerstellung(String outputPath, String lohn, boolean isErsetzenSelected) {
        try {
            String resourceFilePath = "/Stundenzettel_Vorlage.xlsx";
            InputStream resourceStream = AbstractExcelWriter.class.getResourceAsStream(resourceFilePath);

            if (resourceStream == null) {
                System.err.println("File not found in resources: " + resourceFilePath);
                return;
            }

            // Create the destination path in the home directory
            Path destinationPath = Paths.get(PATH_DATEI_STUNDENZETTELVORLAGE);

            if (!Files.exists(destinationPath)) {
                // Copy the file from resources to the home directory
                Files.copy(resourceStream, destinationPath, StandardCopyOption.REPLACE_EXISTING);
                System.out.println("Stundenzettel_Vorlage wurde kopiert zu" + destinationPath);
            } else {
                System.out.println("Vorlage existiert bereits im Verzeichnis: " + PATH_DATEI_STUNDENZETTELVORLAGE);
            }


            InputStream inputStream = new FileInputStream(PATH_DATEI_STUNDENZETTELVORLAGE);
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

            double svBrutto = Double.parseDouble(this.svBrutto.replace(",", "."));
            double stundenlohn = Double.parseDouble(lohn.replace(",", "."));
            double stundensatz = svBrutto / stundenlohn;

            double meanProportionPerEuro = 2.5 / 520;
            double totalMean;

            if(svBrutto <= 520) {
                totalMean = 2.5;
            } else {
                totalMean = meanProportionPerEuro * svBrutto;
            }

            double arbeitstage = stundensatz / totalMean;
            int gerundeteArbeitstage = (int) Math.round(arbeitstage);
            System.out.println("stundensatz: " + stundensatz);
            if (gerundeteArbeitstage == 0) gerundeteArbeitstage = 1;

            double gerundeterStundensatz = stundensatz * 10;
            gerundeterStundensatz = Math.round(gerundeterStundensatz);
            System.out.println("gerundeter stundensatz: " + gerundeterStundensatz);
            gerundeterStundensatz = gerundeterStundensatz / 10;
            System.out.println("gerundeter stundensatz: " + gerundeterStundensatz);


            // Prüfen, ob die Anzahl der Arbeitstage, die "gearbeitet wurden" auch in den Monat passen
            // Das ist nicht der Fall, wenn bspw. der Stundenlohn im Vergleich zum svBrutto sehr niedrig ist und die Person hätte zu viele Stunden bzw. Tage arbeiten müssen, um das zu erreichen
            if (gerundeteArbeitstage > arbeitszeitenCells.size()) {
                StundenzettelController.displayErrorInGui("Das Gehalt übersteigt die mögliche Monatsarbeitszeit im Verhältnis zum angegebenen Stundenlohn ");
                return;
            }


            //Wir erstellen ein Array mit den Arbeitszeiten
            double[] arbeitszeiten = generateRandomNumbers(gerundeteArbeitstage, totalMean, 1);
            double sum = 0;
            for (double value : arbeitszeiten) {
                sum += value;
            }
            for (int j = 0; j < arbeitszeiten.length; j++) {
                arbeitszeiten[j] = arbeitszeiten[j] * (gerundeterStundensatz / sum);
            }
            double sumAfter = 0;
            for (double value : arbeitszeiten) {
                sumAfter += value;
            }


            String[] werktage = new String[arbeitszeitenCells.size()];
            Random randomNumberGen = new Random();
            int randomNumber;
            DecimalFormat decimalFormat = new DecimalFormat("###.##");
            int tempcounter = 0;


            // Array mit Indices aller werktage wird geshuffelt, die ersten x tage werden nacheinander befüllt, die übrigen hinten im array sind dann die random freien tage
            ArrayList<Integer> listOfIndices = new ArrayList<>();
            for (int i = 0; i < werktage.length; i++) {
                listOfIndices.add(i);
            }
            System.out.println(listOfIndices);
            Collections.shuffle(listOfIndices);
            System.out.println(listOfIndices);

            for (int i = 0; i < arbeitszeiten.length; i++) {
                werktage[listOfIndices.get(i)] = decimalFormat.format(arbeitszeiten[i]);
//                werktage[listOfIndices.get(i)] = String.valueOf(arbeitszeiten[i]);
                System.out.println(Arrays.asList(werktage));
                System.out.println(">> " + tempcounter + " <<");
            }

            double zaahl = 0;
            System.out.println(Arrays.asList(werktage));
            for (String werktag : werktage) {
                if (werktag != null) System.out.println(zaahl += Double.parseDouble(werktag.replace(",", ".")));
            }
            System.out.println(zaahl);
            System.out.println(">> " + tempcounter + " <<");

            try {
                //Wir befüllen die Spalten, Dezimal, Arbeitszeit Netto, Aufgezeichnet am, und Arbeitszeit
                String hourMinutes;
                double insgMinuten, minuten, sekunden;
                int stunden;
                for (int i = 0; i < arbeitszeitenCells.size(); i++) {
                    hourMinutes = "";
                    if (werktage[i] != null) {
                        //Wir befüllen die Spalte "Dezimal"
                        arbeitszeitenCells.get(i).setCellValue(werktage[i]);
                        //Wir befüllen sie Spalte "Arbeitszeit Netto"
                        arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 1).setCellValue(werktage[i]);
                        //Wir befüllen die Spalte "Aufgezeichnet am"
//                    LocalDate aufgezeichnetAm = arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 2).getLocalDateTimeCellValue().toLocalDate().with(TemporalAdjusters.next(DayOfWeek.MONDAY));
                        LocalDate aufgezeichnetAm = arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 2).getLocalDateTimeCellValue().toLocalDate().plusDays(1);
                        DayOfWeek day = aufgezeichnetAm.getDayOfWeek();
                        while (day == DayOfWeek.SUNDAY || isDatumEinFeiertag(aufgezeichnetAm, Integer.parseInt(datum[0]))) {
                            System.out.println("Nächster Tag ist " + day);
                            aufgezeichnetAm = aufgezeichnetAm.plusDays(1);
                            day = aufgezeichnetAm.getDayOfWeek();
                        }
                        arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 2).setCellValue(aufgezeichnetAm);
                        //Wir befüllen die Spalte "Arbeitszeit"
                        String temp = werktage[i].replace(",", ".");
                        insgMinuten = Double.parseDouble(temp) * 60;
                        stunden = (int) Double.parseDouble(temp);
                        minuten = insgMinuten % 60;
                        sekunden = Double.parseDouble("0." + String.valueOf(insgMinuten).split("\\.")[1]);
                        sekunden = sekunden * 60;

                        hourMinutes = hourMinutes + "0" + stunden + ":";
                        if (minuten >= 10) hourMinutes += String.valueOf(minuten).split("\\.")[0];
                        else hourMinutes += "0" + String.valueOf(minuten).split("\\.")[0];
                        if (sekunden >= 10) hourMinutes += ":" + (int) sekunden;
                        else hourMinutes += ":" + "0" + (int) sekunden;


                        arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 1).setCellValue(hourMinutes);

                    } else {

                        arbeitszeitenCells.get(i).setCellValue("");
                        //arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() - 1).setCellValue("");
                        arbeitszeitenCells.get(i).getRow().getCell(arbeitszeitenCells.get(i).getColumnIndex() + 2).setCellValue("");
                    }
                }
            } catch (NumberFormatException e) {
                System.out.println("Fehler");
            }
//            try (FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\MM\\Downloads\\test\\test.xlsx")) {
//                workbook.write(fileOutputStream);
//            }

            String fileName = abrechnungsmonat.replace("/", "-");
            String filePathWithName = outputPath + "\\" + fileName + DOCUMENT_FILE_SUFFIX;
            if (!isErsetzenSelected) {
                File pdfFile = new File(filePathWithName);
                if (pdfFile.exists()) {
//                    Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
//                    alert.setTitle("Gleichnamige Datei gefunden");
//                    alert.setHeaderText("Die Datei \"" + fileName+DOCUMENT_FILE_SUFFIX + "\" existiert bereits in dem Pfad \"" + outputPath + "\". Soll die Datei überschrieben werden?");
//                    alert.getButtonTypes().setAll(ButtonType.YES, ButtonType.NO);
//
//                    alert.showAndWait().ifPresent(response -> {
//                        if (response == ButtonType.YES) {
//                            // User clicked Yes, overwrite the file
//                            PdfGenerator pdfGenerator = new PdfGenerator();
//                            pdfGenerator.createPdf(workbook, outputPath, fileName);
//                            System.out.println("PDF created (overwrite)");
//                        } else {
//                            if (response == ButtonType.NO) {
//                                // if user presses No, nothing happens
//                            }
//                        }
//                    });

                    int count = 0;

                    File newFile;
                    do {
                        count++;
                        String newFileName = outputPath + "\\" + fileName + "_" + count + DOCUMENT_FILE_SUFFIX;
                        newFile = new File(newFileName);
                    } while (newFile.exists());

                    PdfGenerator pdfGenerator = new PdfGenerator();
                    pdfGenerator.createPdf(workbook, outputPath, fileName + "_" + count);
                    System.out.println("PDF file created: " + newFile.getAbsolutePath());
                } else {
                    PdfGenerator pdfGenerator = new PdfGenerator();
                    pdfGenerator.createPdf(workbook, outputPath, fileName);
                }
            } else {
                PdfGenerator pdfGenerator = new PdfGenerator();
                pdfGenerator.createPdf(workbook, outputPath, fileName);
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
            } while (randomValue < 0.25);

            DecimalFormat decimalFormat = new DecimalFormat("###.#");
            result[i] = Double.parseDouble(decimalFormat.format(randomValue).replace(",", "."));
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

        freierTagStyle.setFillPattern(FillPatternType.BRICKS);
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
