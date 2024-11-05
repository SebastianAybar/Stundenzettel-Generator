package com.example.stundenzettel;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import de.focus_shift.jollyday.core.Holiday;
import de.focus_shift.jollyday.core.HolidayManager;
import de.focus_shift.jollyday.core.ManagerParameters;
import org.apache.poi.ss.usermodel.*;
import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.poi.ss.usermodel.Font;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.WeekFields;
import java.util.*;
import java.time.format.DateTimeFormatter;
import java.util.List;

import static com.example.stundenzettel.Attribute.*;
import static java.util.Locale.GERMANY;


public class AbstractExcelWriter {

    private final String BUNDESLAND = "he";
    private final String outputPath;

    AbstractExcelWriter(String outputPath) {
        this.outputPath = outputPath;
    }

    public void writeToExcel(List<List<MitarbeiterMonat>> jahresliste, double stundenlohn, boolean isErsetzenSelected) {
        int counter = 1;
        for (List<MitarbeiterMonat> monatsliste : jahresliste) {
            Workbook workbook = null;
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

                workbook = WorkbookFactory.create(inputStream);
                Sheet currentSheet;
                int counterSheets = 0;
                for (int i = 0; i < monatsliste.size(); i++) {
                    workbook.cloneSheet(0);
                    currentSheet = workbook.getSheetAt(counterSheets + 1);
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

                    //Wir erstellen ein Array mit den normalverteilten Arbeitszeiten
                    double svBrutto = Double.parseDouble(monatsliste.get(i).getSvBrutto().replace(",", "."));

                    if (svBrutto > 1600) {
                        StundenzettelController.displayErrorInGui("Brutto von " + monatsliste.get(i).getNachnameVorname() + " war zu hoch.\nStundenzettel wurde übersprungen.");
                        workbook.removeSheetAt(counterSheets + 1);
                        continue;
                    } else {
                        counterSheets++;
                    }

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
                    int tempcounter = 0;

                    // Array mit Indices aller werktage wird geshuffelt, die ersten x tage werden nacheinander befüllt, die übrigen hinten im array sind dann die random freien tage
                    ArrayList<Integer> listOfIndices = new ArrayList<>();
                    for (int j = 0; j < werktage.length; j++) {
                        listOfIndices.add(j);
                    }
                    System.out.println(listOfIndices);
                    Collections.shuffle(listOfIndices);
                    System.out.println(listOfIndices);

                    for (int j = 0; j < arbeitszeiten.length; j++) {
                        werktage[listOfIndices.get(j)] = decimalFormat.format(Math.floor(arbeitszeiten[j] * 100) / 100);
//                werktage[listOfIndices.get(i)] = String.valueOf(arbeitszeiten[i]);
                        System.out.println(Arrays.asList(werktage));
                        System.out.println(">> " + tempcounter + " <<");
                    }



                    try {
                        //Wir befüllen die Spalten, Dezimal, Arbeitszeit Netto, Aufgezeichnet am, und Arbeitszeit
                        String hourMinutes;
                        double insgMinuten, minuten, sekunden;
                        int stunden;

                        for (int k = 0; k < arbeitszeitenCells.size(); k++) {
                            hourMinutes = "";
                            if (werktage[k] != null) {
                                DecimalFormat df = new DecimalFormat("#.00");
                                double arbeitszeit = Double.parseDouble(werktage[k].replace(",", "."));
                                //Wir befüllen die Spalte "Dezimal"
                                if (arbeitszeit < 6) {
                                    arbeitszeitenCells.get(k).setCellValue(werktage[k]);
                                } else if (arbeitszeit < 9) {
                                    arbeitszeit += 0.5;
                                    String stringArbeitszeit = df.format(arbeitszeit).replace(".", ",");
                                    if (stringArbeitszeit.endsWith("0")) {
                                        stringArbeitszeit = stringArbeitszeit.substring(0, stringArbeitszeit.length() - 1);
                                    }
                                    arbeitszeitenCells.get(k).setCellValue(stringArbeitszeit);
                                } else {
                                    arbeitszeit += 0.75;
                                    String stringArbeitszeit = df.format(arbeitszeit).replace(".", ",");
                                    if (stringArbeitszeit.endsWith("0")) {
                                        stringArbeitszeit = stringArbeitszeit.substring(0, stringArbeitszeit.length() - 1);
                                    }
                                    arbeitszeitenCells.get(k).setCellValue(stringArbeitszeit);;
                                }
                                //Wir befüllen sie Spalte "Arbeitszeit Netto"
                                arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 1).setCellValue(werktage[k]);
                                //Wir befüllen die Spalte "Aufgezeichnet am"
                                LocalDate aufgezeichnetAm = arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 2).getLocalDateTimeCellValue().toLocalDate().plusDays(1);
                                DayOfWeek day = aufgezeichnetAm.getDayOfWeek();
                                while (day == DayOfWeek.SUNDAY || isDatumEinFeiertag(aufgezeichnetAm, Integer.parseInt(datum[0]))) {
                                    System.out.println("Nächster Tag ist " + day);
                                    aufgezeichnetAm = aufgezeichnetAm.plusDays(1);
                                    day = aufgezeichnetAm.getDayOfWeek();
                                }
                                arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 2).setCellValue(aufgezeichnetAm);
                                //Wir befüllen die Spalte "Arbeitszeit"
                                String temp = String.valueOf(arbeitszeit).replace(",", ".");
                                insgMinuten = Double.parseDouble(temp) * 60;
                                stunden = (int) Double.parseDouble(temp);
                                minuten = insgMinuten % 60;
                                sekunden = Double.parseDouble("0." + String.valueOf(insgMinuten).split("\\.")[1]);
                                sekunden = sekunden * 60;

                                if (stunden >= 10) hourMinutes += stunden + ":";
                                else hourMinutes += "0" + stunden + ":";
                                if (minuten >= 10) hourMinutes += String.valueOf(minuten).split("\\.")[0];
                                else hourMinutes += "0" + String.valueOf(minuten).split("\\.")[0];
//                                if (sekunden >= 10) hourMinutes += ":" + (int) sekunden;
//                                else hourMinutes += ":" + "0" + (int) sekunden;

                                arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 1).setCellValue(hourMinutes);
                            } else {
                                arbeitszeitenCells.get(k).setCellValue("");
//                            arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() - 1).setCellValue("");
                                arbeitszeitenCells.get(k).getRow().getCell(arbeitszeitenCells.get(k).getColumnIndex() + 2).setCellValue("");
                            }
                        }
                    } catch (NumberFormatException e) {
                        System.out.println("Error");
                    }

                    //Excel Formeln werden nach dem Füllen der Felder noch einmal ausgeführt (z.B. für KW)
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    evaluator.evaluateAll();

                }

                workbook.removeSheetAt(0);

//                 Excel-Output-Dateien
//                try (FileOutputStream fileOutputStream = new FileOutputStream(outputPath + "\\test" + counter++ + ".xlsx")) {
//                    workbook.write(fileOutputStream);
//                }



                String fileName = monatsliste.get(0).getAbrechnungsmonat().replace("/", "-");
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

////              PDF-Output-Dateien
//                PdfGenerator pdfGenerator = new PdfGenerator();
//                pdfGenerator.createPdf(workbook, outputPath, monatsliste.get(0).getAbrechnungsmonat().replace("/", "-"));




            } catch (Exception e) {
                System.out.println("EXCEPTION im Writer!");
                e.printStackTrace();
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
            } while (randomValue < 0.25);
            DecimalFormat decimalFormat = new DecimalFormat("###.#");
            result[i] = Double.parseDouble(decimalFormat.format(randomValue).replace(",", "."));
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

    // Prüft, ob das angegebene Datum ein Feiertag ist (mithilfe der Library <jollyday>)
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
