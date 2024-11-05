package com.example.stundenzettel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;

import static com.example.stundenzettel.Attribute.*;

public class PdfGenerator {

    public void createPdf(Workbook workbook, String outputPath, String filename) {
        try {

            // PDF-Dokument erstellen und öffnen
            Document document = new Document(PAGE_DIN_SIZE, PAGE_MARGIN_LEFT, PAGE_MARGIN_RIGHT, PAGE_MARGIN_TOP, PAGE_MARGIN_BOTTOM);
            PdfWriter.getInstance(document, new FileOutputStream(outputPath + "\\" + filename + DOCUMENT_FILE_SUFFIX));
            document.open();

            // Iteration durch alle Sheets des Excel-Workbooks
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                document.newPage();

                // Titel erstellen
                Paragraph title = createTitle(TITLE_NAME);

                // Unternehmenslogo erstellen
                try {
                    // logo removed
                    InputStream resourceStream = AbstractExcelWriter.class.getResourceAsStream(resourceFilePath);

                    if (resourceStream == null) {
                        System.err.println("Image file not found in resources: " + resourceFilePath);
                        return;
                    }

                    // Create the destination path in the home directory
                    Path destinationPath = Paths.get(PATH_DATEI_LOGO);

                    if (!Files.exists(destinationPath)) {
                        // Copy the file from resources to the home directory
                        Files.copy(resourceStream, destinationPath, StandardCopyOption.REPLACE_EXISTING);
                        System.out.println("Logo wurde kopiert zu " + destinationPath);
                    } else {
                        System.out.println("Logo existiert bereits im Verzeichnis: " + PATH_DATEI_LOGO);
                    }
                } catch (Exception e) {
                    System.out.println("Logo problem hier");
                }
                Image logo = createLogo(PATH_DATEI_LOGO);

                // Mitarbeiterinfotabelle (Mitarbeitername, Mitarbeiternummer, Jahr/Monat) erstellen
                PdfPTable mitarbeiterinfoTabelle = createMitarbeiterinfoTabelle(workbook.getSheetAt(i));

                // Stundentabelle erstellen
                PdfPTable stundenTabelle = createStundenTabelle(workbook.getSheetAt(i));

                // Objekte in das PDF-Dokument hinzufügen
                document.add(title);
                document.add(logo);
                document.add(mitarbeiterinfoTabelle);
                document.add(stundenTabelle);
            }

            // PDF-Dokument schließen
            document.close();
            System.out.println("Successful.");
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private Paragraph createTitle(String title) {
        Paragraph paragraph = new Paragraph(title, TITLE_FONT);
        paragraph.setSpacingBefore(TITLE_SPACING_BEFORE);
        paragraph.setSpacingAfter(TITLE_SPACING_AFTER);
        paragraph.setIndentationLeft(TITLE_INDENTATION_LEFT);

        return paragraph;
    }

    private Image createLogo(String path) throws BadElementException, IOException {
        Image logo = Image.getInstance(path);
        logo.scalePercent(LOGO_SCALE_PERCENT);
        logo.setAbsolutePosition(LOGO_ABSOLUTE_POS_X, LOGO_ABSOLUTE_POS_Y);

        return logo;
    }

    private PdfPTable createMitarbeiterinfoTabelle(Sheet sheet) throws DocumentException {
        PdfPTable mitarbeiterinfoTabelle = new PdfPTable(TBL_MITARBEITERINFO_COLUMNS_ANZAHL);
        mitarbeiterinfoTabelle.setWidths(new float[]{TBL_MITARBEITERINFO_COLUMN1_WIDTH, TBL_MITARBEITERINFO_COLUMN2_WIDTH, TBL_MITARBEITERINFO_COLUMN3_WIDTH});
        mitarbeiterinfoTabelle.setWidthPercentage(TBL_MITARBEITERINFO_WIDTH_PERCENT);
        mitarbeiterinfoTabelle.setHorizontalAlignment(Element.ALIGN_LEFT);
        mitarbeiterinfoTabelle.setSpacingAfter(TBL_MITARBEITERINFO_SPACING_AFTER);

        PdfPCell zelle = new PdfPCell();
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_MITARBEITERNAME_TEXT).getCell(COLUMN_MITARBEITERNAME_TEXT).getStringCellValue(), TBL_MITARBEITERINFO_FETT_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_MITARBEITERNAME).getCell(COLUMN_MITARBEITERNAME).getStringCellValue(), TBL_MITARBEITERINFO_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell();
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_MITARBEITERNUMMER_TEXT).getCell(COLUMN_MITARBEITERNUMMER_TEXT).getStringCellValue(), TBL_MITARBEITERINFO_FETT_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_MITARBEITERNUMMER).getCell(COLUMN_MITARBEITERNUMMER).getStringCellValue(), TBL_MITARBEITERINFO_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell();
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_JAHRMONAT_TEXT).getCell(COLUMN_JAHRMONAT_TEXT).getStringCellValue(), TBL_MITARBEITERINFO_FETT_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        zelle = new PdfPCell(new Paragraph(sheet.getRow(ROW_JAHRMONAT).getCell(COLUMN_JAHRMONAT).getStringCellValue(), TBL_MITARBEITERINFO_FONT));
        zelle.setBorder(CELL_BORDER_TRANSPARENT);
        mitarbeiterinfoTabelle.addCell(zelle);

        return mitarbeiterinfoTabelle;
    }

    private PdfPTable createStundenTabelle(Sheet sheet) throws DocumentException {
        PdfPTable stundenzettelTabelle = new PdfPTable(TBL_STUNDEN_COLUMN_ANZAHL);
        float widthsDerSpalten[] = {KW_COLUMN_WIDTH, WOCHENTAG_COLUMN_WIDTH, DATUM_COLUMN_WIDTH, ARBEITSZEIT_COLUMN_WIDTH,
                DEZIMAL_COLUMN_WIDTH, ARBEITSZEITNETTO_COLUMN_WIDTH, AUFGEZEICHNETAM_COLUMN_WIDTH};
        stundenzettelTabelle.setWidths(widthsDerSpalten);
        stundenzettelTabelle.setHeaderRows(TBL_STUNDEN_HEADER_ANZAHL);

        addKopfzeileToStundenTabelle(stundenzettelTabelle, sheet);

        Phrase pdfZelleninhalt = new Phrase(DEFAULT_VALUE);
        PdfPCell pdfZelle = new PdfPCell(pdfZelleninhalt);

        int insgArbeitszeitStunden = 0, insgArbeitszeitMinuten = 0, insgArbeitszeitSekunden = 0;
        String insgArbeitszeit = "";
        double insgDezimal = 0, insgArbeitszeitNetto = 0;


        for (int i = ROWINDEX_ERSTES_DATUM; i < sheet.getLastRowNum() + 1; i++) {
            for (Cell excelZelle : sheet.getRow(i)) {
                if (excelZelle.getCellType() == CellType.STRING) {
                    String zelleAlsString = excelZelle.getStringCellValue();
                    pdfZelleninhalt = new Phrase(zelleAlsString, TBL_STUNDEN_FONT);
                    if (excelZelle.getColumnIndex() == 3) {
                        if (!zelleAlsString.isEmpty()) {
                            int stunde = Integer.parseInt(zelleAlsString.split(":")[0]);
                            int minute = Integer.parseInt(zelleAlsString.split(":")[1]);
                            //int sekunde = Integer.parseInt(zelleAlsString.split(":")[2]);
                            insgArbeitszeitStunden += stunde;
                            insgArbeitszeitMinuten += minute;
//                            insgArbeitszeitSekunden += sekunde;
//
//                            if (insgArbeitszeitSekunden >= 60) {
//                                insgArbeitszeitMinuten += insgArbeitszeitSekunden / 60;
//                                insgArbeitszeitSekunden %= 60;
//                            }

                            if (insgArbeitszeitMinuten >= 60) {
                                insgArbeitszeitStunden += insgArbeitszeitMinuten / 60;
                                insgArbeitszeitMinuten %= 60;
                            }
                        }
                    }
                    if (excelZelle.getColumnIndex() == 4) {
                        if (!zelleAlsString.isEmpty()) {
                            insgDezimal += Double.parseDouble(zelleAlsString.replace(",", "."));
                        }
                    }
                    if (excelZelle.getColumnIndex() == 5) {
                        if (!zelleAlsString.isEmpty()) {
                            insgArbeitszeitNetto += Double.parseDouble(zelleAlsString.replace(",", "."));
                        }
                    }
                } else if (excelZelle.getCellType() == CellType.NUMERIC) {
                    // Wenn Spalte "KW"
                    if (excelZelle.getColumnIndex() == 0) {
                        pdfZelleninhalt = new Phrase(String.valueOf((int) excelZelle.getNumericCellValue()), TBL_STUNDEN_FONT);
                        // Wenn Spalte "Datum" oder "Aufgezeichnet am"
                    } else {
                        LocalDate datum = convertNumericCellToDate(excelZelle);
                        String datumDeutschesFormat = formatGermanDate(datum);
                        pdfZelleninhalt = new Phrase(datumDeutschesFormat, TBL_STUNDEN_FONT);
                    }
                } else {
                    pdfZelleninhalt = new Phrase("", TBL_STUNDEN_FONT);
                }

                pdfZelle = new PdfPCell(pdfZelleninhalt);

                if (excelZelle.getCellStyle().getFillPattern() == FillPatternType.BRICKS) {
                    pdfZelle.setBackgroundColor(FREIERTAG_COLOR);
                }

                pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
                pdfZelle.setPadding(TBL_STUNDEN_PADDING);
                stundenzettelTabelle.addCell(pdfZelle);
            }
        }

        int insgArbeitszeitHours;
        double insgArbeitszeitMinutes = insgDezimal * 60;
        insgArbeitszeitHours = (int) insgDezimal;
        double minuten = insgArbeitszeitMinutes % 60;
        double sekunden = Double.parseDouble("0." + String.valueOf(insgArbeitszeitMinutes).split("\\.")[1]);
        sekunden *= 60;


        if (insgArbeitszeitHours >= 10) insgArbeitszeit += insgArbeitszeitHours + ":";
        else insgArbeitszeit += "0" + insgArbeitszeitHours + ":";
        if (minuten >= 10) insgArbeitszeit += (int) minuten;
        else insgArbeitszeit += "0" + (int) minuten;
//        if (insgArbeitszeitSekunden >= 10) insgArbeitszeit += insgArbeitszeitSekunden;
//        else insgArbeitszeit += "0" + insgArbeitszeitSekunden;

        addSummenzeileToStundenTabelle(stundenzettelTabelle, insgArbeitszeit, insgDezimal, insgArbeitszeitNetto);

        return stundenzettelTabelle;
    }

    private void addKopfzeileToStundenTabelle(PdfPTable stundenTabelle, Sheet sheet) {
        String kopfzeileKw = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_KW).getStringCellValue();
        PdfPCell pdfZelle = new PdfPCell(new Phrase(kopfzeileKw, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileWochentag = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_WOCHENTAG).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileWochentag, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileDatum = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_DATUM).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileDatum, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileArbeitszeit = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_ARBEITSZEIT).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileArbeitszeit, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileDezimal = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_DEZIMAL).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileDezimal, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileArbeitszeitNetto = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_ARBEITSZEITNETTO).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileArbeitszeitNetto, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);

        String kopfzeileAufgezeichnetAm = sheet.getRow(ROW_KOPFZEILE).getCell(COLUMN_AUFGEZEICHNETAM).getStringCellValue();
        pdfZelle = new PdfPCell(new Phrase(kopfzeileAufgezeichnetAm, KOPFZEILE_FONT));
        pdfZelle.setBackgroundColor(KOPFZEILE_COLOR);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        stundenTabelle.addCell(pdfZelle);
    }

    private void addSummenzeileToStundenTabelle(PdfPTable stundenTabelle, String insgArbeitszeit, double insgDezimal, double insgArbeitszeitNetto) {
        // Letzte Zeile: leer, leer, leer, Arbeitszeitsumme, Dezimalsumme, ArbeitszeitNettosumme, leer
        DecimalFormat insgSummenFormat = new DecimalFormat(TBL_STUNDEN_SUMMEN_FORMAT);
        PdfPCell pdfZelle = new PdfPCell();

        // Leere Zelle
        stundenTabelle.addCell(pdfZelle);

        // Leere Zelle
        stundenTabelle.addCell(pdfZelle);

        // Leere Zelle
        stundenTabelle.addCell(pdfZelle);

        // Zelle: Arbeitszeit
        Phrase zelleninhalt = new Phrase(insgArbeitszeit, TBL_STUNDEN_FETT_FONT);
        pdfZelle = new PdfPCell(zelleninhalt);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        pdfZelle.setPadding(TBL_STUNDEN_PADDING);
        stundenTabelle.addCell("");
//        stundenTabelle.addCell(pdfZelle);

        // Zelle: Dezimal
        zelleninhalt = new Phrase(insgSummenFormat.format(insgDezimal), TBL_STUNDEN_FETT_FONT);
        pdfZelle = new PdfPCell(zelleninhalt);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        pdfZelle.setPadding(TBL_STUNDEN_PADDING);
        stundenTabelle.addCell(pdfZelle);

        // Zelle: Arbeitszeit netto
        zelleninhalt = new Phrase(String.valueOf(insgSummenFormat.format(insgArbeitszeitNetto)), TBL_STUNDEN_FETT_FONT);
        pdfZelle = new PdfPCell(zelleninhalt);
        pdfZelle.setHorizontalAlignment(Element.ALIGN_CENTER);
        pdfZelle.setPadding(TBL_STUNDEN_PADDING);
        stundenTabelle.addCell(pdfZelle);

        pdfZelle = new PdfPCell();
        stundenTabelle.addCell(pdfZelle);
    }


    public static LocalDate convertNumericCellToDate(Cell cell) {
        if (DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        } else {
            DataFormatter dataFormatter = new DataFormatter();
            String cellValueAsString = dataFormatter.formatCellValue(cell);
            double numericCellValue = Double.parseDouble(cellValueAsString);
            return LocalDate.ofEpochDay((long) numericCellValue);
        }
    }

    public static String formatGermanDate(LocalDate localDate) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_GERMAN_FORMAT);
        return localDate.format(formatter);
    }
}