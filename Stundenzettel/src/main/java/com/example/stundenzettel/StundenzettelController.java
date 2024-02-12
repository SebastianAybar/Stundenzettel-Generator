package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.*;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.net.URL;
import java.util.List;
import java.util.ResourceBundle;

import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;

import static com.example.stundenzettel.Attribute.*;

public class StundenzettelController implements Initializable {

    private AbstractExcelReader abstractExcelReader;
    private AbstractExcelWriter abstractExcelWriter;
    private EinzelerstellungReader einzelerstellungReader;
    private boolean isExcelListeClicked = true;
    private boolean isEinzelerstellungClicked = false;

    @FXML
    private Label lblStundenlohn;
    @FXML
    private Label lblDatei;
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
    private Label lblFalschesFormatStundenlohn;
    @FXML
    private Label lblDateiNichtAkzeptiert;
    @FXML
    private Label lblFalscherPathOutput;

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
//        inputPathTextField.setText("C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Documents\\Mini-Job geringfügig Beschäftigte_01_10_2023_LV_Testnamen.xlsx");
//        outputPathTextField.setText("C:\\Users\\sebas\\OneDrive\\Dokumente\\GitHub\\Stundenzettel-Generator\\Documents");
//        inputPathTextField.setText("C:\\Users\\MM\\Downloads\\ifi_USB-Stick\\bsp.xlsx");
//        inputPathTextField.setText("C:\\Users\\MM\\Downloads\\ifi_USB-Stick\\einEintrag.xlsx");
//        inputPathTextField.setText("C:\\Users\\MM\\Downloads\\ifi_USB-Stick\\einMonat.xlsx");
//        inputPathTextField.setText("C:\\Users\\MM\\Downloads\\ifi_USB-Stick\\mitLuecke.xlsx");
//        inputPathTextField.setText("C:\\Users\\MM\\Downloads\\ifi_USB-Stick\\komplettFalschesFormat.xlsx");

//        outputPathTextField.setText("C:\\Users\\MM\\Downloads\\test");

        if (isExcelListeClicked) {
            boolean isFeldStundenlohnGueltig = false;
            boolean isFeldPathInputGueltig = false;
            boolean isFeldPathOutputGueltig = false;

            // Prüfung: Feld Stundenlohn
            isFeldStundenlohnGueltig = checkFieldStundenlohn();

            // Prüfung: Feld PathInput
            if(!inputPathTextField.getText().isEmpty()) {
                defaultPathInput();
                isFeldPathInputGueltig = true;
            } else {
                falscherPathInput();
                isFeldPathInputGueltig = false;
            }

            // Prüfung: Feld PathOutput
            if(!outputPathTextField.getText().isEmpty()) {
                defaultPathOutput();
                isFeldPathOutputGueltig = true;
            } else {
                falscherPathOutput();
                isFeldPathOutputGueltig = false;
            }

            if (isFeldStundenlohnGueltig && isFeldPathInputGueltig && isFeldPathOutputGueltig) {
                saveStundenlohnToDatei(textFieldStundenlohn.getText());

                // Prüfung: Feld PathInput: try/catch ist gleichzeitig Prüfung für inputPathField (richtiges Dateiformat oder nicht)
                try {
                    abstractExcelReader = new AbstractExcelReader(inputPathTextField.getText());
                    List<List<MitarbeiterMonat>> jahresliste = abstractExcelReader.getListOfAbrechnungsmonate();
                    defaultFormatExcelDatei();

                    abstractExcelWriter = new AbstractExcelWriter(outputPathTextField.getText());
                    abstractExcelWriter.writeToExcel(jahresliste, Double.parseDouble(textFieldStundenlohn.getText()));
                } catch (Exception e) {
                    System.out.println("Falsches Format der Excel-Datei");
                    falschesFormatExcelDatei();
                }
            }

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
            boolean isFeldStundenlohnGueltig = false;

            // Prüfung: Feld Mitarbeiternummer
            if (!textFieldMitarbeiternummer.getText().isEmpty()) {
                defaultMitarbeiternummer();
                isFeldMitarbeiternummerGueltig = true;
            } else {
                mitarbeiternummerEmpty();
                isFeldMitarbeiternummerGueltig = false;
            }

            // Prüfung: Feld Name
            if (!textFieldName.getText().isEmpty()) {
                defaultName();
                isFeldNameGueltig = true;
            } else {
                nameEmpty();
                isFeldNameGueltig = false;
            }

            // Prüfung: Feld Abrechnungsmonat
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

            // Prüfung: Feld SvBrutto
            try {
                double svBrutto = Double.parseDouble(textFieldSvBrutto.getText().replace(",", "."));
                if (svBrutto >= 0) {
                    defaultFormatSvBrutto();
                    isFeldSvBruttoGueltig = true;
                } else {
                    falschesFormatSvBrutto();
                    isFeldSvBruttoGueltig = false;
                }
            } catch (NumberFormatException nfe) {
                falschesFormatSvBrutto();
                isFeldSvBruttoGueltig = false;
            } catch (NullPointerException npe) {
                falschesFormatSvBrutto();
                isFeldSvBruttoGueltig = false;
            }

            // Prüfung: Feld Stundenlohn
            isFeldStundenlohnGueltig = checkFieldStundenlohn();


            if (isFeldAbrechnungsmonatGueltig && isFeldSvBruttoGueltig && isFeldNameGueltig && isFeldMitarbeiternummerGueltig && isFeldStundenlohnGueltig) {
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


    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {

        loadStundenlohnFromDatei();
    }


    private void loadStundenlohnFromDatei() {
        String pathStundenlohnDatei = "E:\\zAndere\\GitRepos\\Stundenzettel-Generator\\Stundenzettel\\src\\main\\resources\\com\\example\\stundenzettel\\stundenlohn.txt";
        String defaultStundenlohn = DEFAULT_STUNDENLOHN;

        try (BufferedReader reader = new BufferedReader(new FileReader(pathStundenlohnDatei))) {
            String zeile = reader.readLine();
            if (zeile != null && !zeile.isEmpty()) {
                textFieldStundenlohn.setText(zeile.trim());
            }
        } catch (IOException e) {
            e.printStackTrace();
            textFieldStundenlohn.setText(defaultStundenlohn);
        }
    }

    private void saveStundenlohnToDatei(String stundenlohn) {
        try (FileWriter fileWriter = new FileWriter("E:\\zAndere\\GitRepos\\Stundenzettel-Generator\\Stundenzettel\\src\\main\\resources\\com\\example\\stundenzettel\\stundenlohn.txt")) {
            fileWriter.write(stundenlohn);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private boolean checkFieldStundenlohn() {
        if (!textFieldStundenlohn.getText().isEmpty()) {
            try {
                double stundenlohn = Double.parseDouble(textFieldStundenlohn.getText().replace(",", "."));

                if (stundenlohn >= 0) {
                    defaultFormatStundenlohn();
                    return true;
                } else {
                    falschesFormatStundenlohn();
                    return false;
                }

            } catch (NumberFormatException e) {
                e.printStackTrace();
                falschesFormatStundenlohn();
                return false;
            } catch (NullPointerException e) {
                e.printStackTrace();
                falschesFormatStundenlohn();
                return false;
            }
        } else {
            falschesFormatStundenlohn();
            return false;
        }
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

    protected void falschesFormatStundenlohn() {
        textFieldStundenlohn.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        textFieldStundenlohn.setBorder(border);
        lblFalschesFormatStundenlohn.setVisible(true);
    }

    protected void defaultFormatStundenlohn() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldStundenlohn.setBorder(border);
        lblFalschesFormatStundenlohn.setVisible(false);
    }

    protected void falschesFormatExcelDatei() {
        inputPathTextField.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        inputPathTextField.setBorder(border);
        lblDateiNichtAkzeptiert.setText("Die Inhalte der Datei konnten teilweise oder gar nicht gelesen werden");
        lblDateiNichtAkzeptiert.setVisible(true);
    }

    protected void defaultFormatExcelDatei() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        inputPathTextField.setBorder(border);
        lblDateiNichtAkzeptiert.setVisible(false);
    }

    protected void falscherPathInput() {
        inputPathTextField.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        inputPathTextField.setBorder(border);
        lblDateiNichtAkzeptiert.setText("Ungültige Eingabe");
        lblDateiNichtAkzeptiert.setVisible(true);
    }

    protected void defaultPathInput() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        inputPathTextField.setBorder(border);
        lblDateiNichtAkzeptiert.setVisible(false);
    }

    protected void falscherPathOutput() {
        outputPathTextField.setText("");
        Border border = new Border(new BorderStroke(Color.RED, BorderStrokeStyle.SOLID, null, new BorderWidths(2)));
        outputPathTextField.setBorder(border);
        lblFalscherPathOutput.setVisible(true);
    }

    protected void defaultPathOutput() {
        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        outputPathTextField.setBorder(border);
        lblFalscherPathOutput.setVisible(false);
    }

}
