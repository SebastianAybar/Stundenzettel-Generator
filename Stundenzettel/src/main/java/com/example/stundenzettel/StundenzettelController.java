package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;

import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;

import static java.util.Collections.*;

public class StundenzettelController {

    private AbstractExcelReader abstractExcelReader;
    private AbstractExcelWriter abstractExcelWriter;
    private EinzelerstellungReader einzelerstellungReader;
    private boolean isExcelListeClicked = true;
    private boolean isEinzelerstellungClicked = false;


    @FXML
    private Label lblStundenlohn;
    @FXML
    private TextField textFieldStundenlohn;
    @FXML
    private AnchorPane anchorPane;
    @FXML
    private Button excelList;
    @FXML
    private Button einzelerstellung;
    @FXML
    private Button btnChooseInputFile;
    @FXML
    private Button btnChooseOutputFile;
    @FXML
    private Label inputFile;
    @FXML
    private Label outputFile;
    @FXML
    private TextField inputPathTextField;
    @FXML
    private TextField outputPathTextField;
    @FXML
    private CheckBox replaceFile;
    @FXML
    private Separator separator1;
    @FXML
    private Separator separator2;
    @FXML
    private Button calculate;
    @FXML
    private FontAwesomeIconView icnChooseInputFile;
    @FXML
    private FontAwesomeIconView icnChooseOutputFile;
    @FXML
    private HBox hboxExcelListeAnsicht;
    @FXML
    private HBox hboxEinzelerstellung;
    @FXML
    private TextField textFieldSvBrutto;
    @FXML
    private TextField textFieldName;
    @FXML
    private TextField textFieldMitarbeiternummer;
    @FXML
    private TextField textFieldAbrechnungsmonat;
    @FXML
    private Label lblAbrechnungsmonat;
    @FXML
    private Label lblMitarbeiternummer;
    @FXML
    private Label lblName;
    @FXML
    private Label lblSvBrutto;
    @FXML
    private Label lblFalschesFormatAbrechnungsmonat;
    @FXML
    private Label lblFalschesFormatSvBrutto;
    @FXML
    private Label lblMitarbeiternummerEmpty;
    @FXML
    private Label lblNameEmpty;

    @FXML
    protected void chooseFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Datei auswählen");
        File selectedFile = fileChooser.showOpenDialog(null);

        if (selectedFile != null) {
            inputPathTextField.setText(selectedFile.getAbsolutePath());
        }
    }

    @FXML
    protected void chooseDirectory() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Ordner auswählen");
        File selectedDirectory = directoryChooser.showDialog(null);

        if (selectedDirectory != null) {
            outputPathTextField.setText(selectedDirectory.getAbsolutePath());
        }
    }

    @FXML
    protected void transformExcel() {

        inputPathTextField.setText("C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Documents\\Mini-Job geringfügig Beschäftigte_01_10_2023_LV_Testnamen.xlsx");
        outputPathTextField.setText("C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Documents");

        if (isExcelListeClicked) {
            abstractExcelReader = new AbstractExcelReader(inputPathTextField.getText());
            List<List<MitarbeiterMonat>> jahresliste = abstractExcelReader.getListOfAbrechnungsmonate();
            abstractExcelWriter = new AbstractExcelWriter(outputPathTextField.getText());
            abstractExcelWriter.writeToExcel(jahresliste);
        } else if (isEinzelerstellungClicked) {

            /*
            * Prüfung für Feld 1 (bool richtig oder falsch)
            * Prüfung für Feld 2 (bool richtig oder falsch)
            * Prüfung für Feld 3 (bool richtig oder falsch)
            * Prüfung für Feld 4 (bool richtig oder falsch)
            *
            * if(alle bools = true)
            *   Excel erstellen
            *
            * */

            boolean isFeldAbrechnungsmonatGueltig = false;
            boolean isFeldSvBruttoGueltig = false;
            boolean isFeldMitarbeiternummerGueltig = false;
            boolean isFeldNameGueltig = false;

            if (!textFieldMitarbeiternummer.getText().isEmpty()) {
                defaultMitarbeiternummer();
                isFeldMitarbeiternummerGueltig = true;
            } else {
                mitarbeiternummerEmpty();
                isFeldMitarbeiternummerGueltig = false;
            }

            if (!textFieldName.getText().isEmpty()) {
                defaultName();
                isFeldNameGueltig = true;
            } else {
                nameEmpty();
                isFeldNameGueltig = false;
            }

            if (isValidDateFormat(textFieldAbrechnungsmonat.getText())) {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM");
                YearMonth abrechnungsmonat = YearMonth.parse(textFieldAbrechnungsmonat.getText(), formatter);
                if (abrechnungsmonat.getYear() <= LocalDate.now().getYear()) {
                    defaultFormatAbrechnungsmonat();
                    isFeldAbrechnungsmonatGueltig = true;
                } else {
                    falschesFormatAbrechnungsmonat();
                    isFeldAbrechnungsmonatGueltig = false;
                }
            } else {
                falschesFormatAbrechnungsmonat();
                isFeldAbrechnungsmonatGueltig = false;
            }


            try {
                double svBrutto = Double.parseDouble(textFieldSvBrutto.getText().replace(",", "."));
                if (svBrutto >= 0) {
                    defaultFormatSvBrutto();
                    isFeldSvBruttoGueltig = true;
                }

            } catch (NumberFormatException nfe) {
                falschesFormatSvBrutto();
                isFeldSvBruttoGueltig = false;
            } catch (NullPointerException npe) {
                falschesFormatSvBrutto();
                isFeldSvBruttoGueltig = false;
            }

            if (isFeldAbrechnungsmonatGueltig && isFeldSvBruttoGueltig && isFeldNameGueltig && isFeldMitarbeiternummerGueltig) {
                einzelerstellungReader = new EinzelerstellungReader(textFieldAbrechnungsmonat.getText(), textFieldMitarbeiternummer.getText(), textFieldSvBrutto.getText(), textFieldName.getText());
                einzelerstellungReader.writeToExcelEinzelerstellung();
            }
        }
    }

    @FXML
    protected void einzerstellung() {
        hboxExcelListeAnsicht.setVisible(false);
        hboxEinzelerstellung.setVisible(true);

        isEinzelerstellungClicked = true;
        isExcelListeClicked = false;
    }

    @FXML
    protected void excelList() {
        hboxExcelListeAnsicht.setVisible(true);
        hboxEinzelerstellung.setVisible(false);

        isExcelListeClicked = true;
        isEinzelerstellungClicked = false;

        //Wenn Excel Liste angeklickt wird sollen alle Einträge aus Einzelerstellung entfernt werden
        textFieldAbrechnungsmonat.setText("");
        textFieldSvBrutto.setText("");
        textFieldMitarbeiternummer.setText("");
        textFieldName.setText("");
        lblNameEmpty.setVisible(false);
        lblMitarbeiternummerEmpty.setVisible(false);
        lblFalschesFormatSvBrutto.setVisible(false);
        lblFalschesFormatAbrechnungsmonat.setVisible(false);
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldAbrechnungsmonat.setBorder(border);
        textFieldAbrechnungsmonat.setPromptText("Format: yyyy/MM");
        textFieldSvBrutto.setBorder(border);
        textFieldMitarbeiternummer.setBorder(border);
        textFieldName.setBorder(border);

    }

    protected boolean isValidDateFormat(String text) {
        return text.matches("\\d{4}/\\d{2}");
    }

    protected void falschesFormatAbrechnungsmonat() {
        textFieldAbrechnungsmonat.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        textFieldAbrechnungsmonat.setBorder(border);
        lblFalschesFormatAbrechnungsmonat.setVisible(true);
    }

    protected void defaultFormatAbrechnungsmonat() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldAbrechnungsmonat.setBorder(border);
        lblFalschesFormatAbrechnungsmonat.setVisible(false);
    }

    protected void falschesFormatSvBrutto() {
        textFieldSvBrutto.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        textFieldSvBrutto.setBorder(border);
        lblFalschesFormatSvBrutto.setVisible(true);
    }

    protected void defaultFormatSvBrutto() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldSvBrutto.setBorder(border);
        lblFalschesFormatSvBrutto.setVisible(false);
    }

    protected void mitarbeiternummerEmpty() {
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        textFieldMitarbeiternummer.setBorder(border);
        lblMitarbeiternummerEmpty.setVisible(true);
    }

    protected void defaultMitarbeiternummer() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldMitarbeiternummer.setBorder(border);
        lblMitarbeiternummerEmpty.setVisible(false);
    }

    protected void nameEmpty() {
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        textFieldName.setBorder(border);
        lblNameEmpty.setVisible(true);
    }

    protected void defaultName() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldName.setBorder(border);
        lblNameEmpty.setVisible(false);
    }

}

