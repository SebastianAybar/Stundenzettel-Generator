package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
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
    private Label lblSchlussnachricht;
    @FXML
    private CheckBox checkboxErsetzen;

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

        if (isExcelListeClicked) {
            defaultSchlussnachricht();
            boolean isFeldStundenlohnGueltig = false;
            boolean isFeldPathInputGueltig = false;
            boolean isFeldPathOutputGueltig = false;

            // Prüfung: Feld Stundenlohn
            isFeldStundenlohnGueltig = checkFieldStundenlohn();

            // Prüfung: Feld PathInput
            if (!inputPathTextField.getText().isEmpty()) {
                defaultPathInput();
                isFeldPathInputGueltig = true;
            } else {
                falscherPathInput();
            }

            // Prüfung: Feld PathOutput
            if (!outputPathTextField.getText().isEmpty()) {
                defaultPathOutput();
                isFeldPathOutputGueltig = true;
            } else {
                falscherPathOutput();
            }

            if (isFeldStundenlohnGueltig && isFeldPathInputGueltig && isFeldPathOutputGueltig) {
                saveStundenlohnToDatei(textFieldStundenlohn.getText());

                // Prüfung: Feld PathInput: try/catch ist gleichzeitig Prüfung für inputPathField (richtiges Dateiformat oder nicht)
                File directory = new File(outputPathTextField.getText());
                if (directory.exists() && directory.isDirectory()) {
                    try {
                        abstractExcelReader = new AbstractExcelReader(inputPathTextField.getText());
                        List<List<MitarbeiterMonat>> jahresliste = abstractExcelReader.getListOfAbrechnungsmonate(textFieldStundenlohn.getText().replace(",", "."));
                        defaultFormatExcelDatei();

                        if (jahresliste.isEmpty()) {
                            fehlgeschlageneSchlussnachricht();
                            return;
                        }
                        abstractExcelWriter = new AbstractExcelWriter(outputPathTextField.getText());
                        abstractExcelWriter.writeToExcel(jahresliste, Double.parseDouble(textFieldStundenlohn.getText().replace(",", ".")));
                        erfolgreicheSchlussnachricht();
                    } catch (Exception e) {
//                    e.printStackTrace();
                        System.out.println("Falsches Format der Excel-Datei");
                        falschesFormatExcelDatei();
                    }
                } else {
                    System.out.println("Ist kein directory");
                    falscherPathOutput();
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

            defaultSchlussnachricht();

            boolean isFeldAbrechnungsmonatGueltig = false;
            boolean isFeldSvBruttoGueltig = false;
            boolean isFeldMitarbeiternummerGueltig = false;
            boolean isFeldNameGueltig = false;
            boolean isFeldStundenlohnGueltig = false;
            boolean isFeldPathOutputGueltig = false;

            // Prüfung: Feld Stundenlohn
            isFeldStundenlohnGueltig = checkFieldStundenlohn();

            // Prüfung: Feld PathOutput
            if (!outputPathTextField.getText().isEmpty()) {
                defaultPathOutput();
                isFeldPathOutputGueltig = true;
            } else {
                falscherPathOutput();
            }

            // Prüfung: Feld Mitarbeiternummer
            if (!textFieldMitarbeiternummer.getText().isEmpty()) {
                defaultMitarbeiternummer();
                isFeldMitarbeiternummerGueltig = true;
            } else {
                mitarbeiternummerEmpty();
            }

            // Prüfung: Feld Name
            if (!textFieldName.getText().isEmpty()) {
                defaultName();
                isFeldNameGueltig = true;
            } else {
                nameEmpty();
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
                }
            } else {
                falschesFormatAbrechnungsmonat();
            }

            // Prüfung: Feld SvBrutto
            try {
                double svBrutto = Double.parseDouble(textFieldSvBrutto.getText().replace(",", "."));
                if (svBrutto >= 0) {
                    defaultFormatSvBrutto();
                    isFeldSvBruttoGueltig = true;
                } else {
                    falschesFormatSvBrutto();
                }
            } catch (NumberFormatException | NullPointerException nfe) {
                falschesFormatSvBrutto();
            }


            if (isFeldAbrechnungsmonatGueltig && isFeldSvBruttoGueltig && isFeldNameGueltig && isFeldMitarbeiternummerGueltig && isFeldStundenlohnGueltig) {
                saveStundenlohnToDatei(textFieldStundenlohn.getText());
                if (Double.parseDouble(textFieldSvBrutto.getText().replace(",", ".")) >= Double.parseDouble(textFieldStundenlohn.getText().replace(",", "."))) {
                    File directory = new File(outputPathTextField.getText());
                    if (directory.exists() && directory.isDirectory()) {
                        isErrorDisplayed = false;
                        einzelerstellungReader = new EinzelerstellungReader(textFieldAbrechnungsmonat.getText(), textFieldMitarbeiternummer.getText(), textFieldSvBrutto.getText(), textFieldName.getText());
                        einzelerstellungReader.writeToExcelEinzelerstellung(outputPathTextField.getText(), textFieldStundenlohn.getText(), checkboxErsetzen.isSelected());
                        if(!isErrorDisplayed) erfolgreicheSchlussnachricht();
                    } else {
                        System.out.println("Ist kein directory");
                        falscherPathOutput();
                    }
                } else {
                    falschesFormatSvBrutto();
                    fehlgeschlageneSchlussnachricht();
                }
            }

        }
    }

    @FXML
    protected void einzelerstellung() {
        hboxExcelListeAnsicht.setVisible(false);
        hboxEinzelerstellung.setVisible(true);

        isEinzelerstellungClicked = true;
        isExcelListeClicked = false;

        defaultPathInput();
        defaultPathOutput();

        inputPathTextField.setText("");

        /*lblFalscherPathOutput.setVisible(false);
        lblDateiNichtAkzeptiert.setVisible(false);*/

        /*Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        inputPathTextField.setBorder(border);
        inputPathTextField.setPromptText("Hier können Sie den Pfad der Excel-Liste eingeben.");
        outputPathTextField.setBorder(border);*/
    }

    @FXML
    protected void excelList() {
        hboxExcelListeAnsicht.setVisible(true);
        hboxEinzelerstellung.setVisible(false);

        isExcelListeClicked = true;
        isEinzelerstellungClicked = false;

        defaultPathOutput();

        defaultFormatAbrechnungsmonat();
        defaultFormatSvBrutto();
        defaultName();
        defaultMitarbeiternummer();

        textFieldMitarbeiternummer.setText("");
        textFieldAbrechnungsmonat.setText("");
        textFieldName.setText("");
        textFieldSvBrutto.setText("");

        /*textFieldSvBrutto.setText("");
        textFieldMitarbeiternummer.setText("");
        textFieldName.setText("");
        lblNameEmpty.setVisible(false);
        lblMitarbeiternummerEmpty.setVisible(false);
        lblFalschesFormatSvBrutto.setVisible(false);
        lblFalschesFormatAbrechnungsmonat.setVisible(false);

        Border border = new Border(new BorderStroke(Color.BLACK, BorderStrokeStyle.NONE, null, new BorderWidths(1)));
        textFieldAbrechnungsmonat.setText("");
        textFieldAbrechnungsmonat.setBorder(border);
        textFieldAbrechnungsmonat.setPromptText("Format: yyyy/MM");

        textFieldSvBrutto.setBorder(border);
        textFieldMitarbeiternummer.setBorder(border);
        textFieldName.setBorder(border);*/
    }


    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        loadStundenlohnFromDatei();
    }


    private void loadStundenlohnFromDatei() {
        Path path = Paths.get(PATH_DATEI_STUNDENLOHN);
        String stundenlohn = DEFAULT_STUNDENLOHN;

        // Wenn stundenlohn.txt im Home-Verzeichnis nicht existiert
        if (!Files.exists(path)) {
            // Neue stundenlohn.txt im Home-Verzeichnis erstellen. Wenns nicht klappt, dann ohne Datei erstellen weiter
            System.out.println("Datei existiert nicht!");
            try {
                Files.createFile(path);
                System.out.println("Neue Datei erstellt in: " + PATH_DATEI_STUNDENLOHN);
            } catch (Exception e) {
                System.out.println("Stundenlohn Datei konnte nicht im Verzeichnis " + PATH_DATEI_STUNDENLOHN + " erstellt werden");
            }

            try (PrintWriter printWriter = new PrintWriter(new FileOutputStream(PATH_DATEI_STUNDENLOHN))) {
                printWriter.println(stundenlohn);
                System.out.println("Wert >" + stundenlohn + "< in die neue Datei geschrieben");
            } catch (IOException e) {
                System.out.println("In die Stundenlohn Datei im Verzeichnis " + PATH_DATEI_STUNDENLOHN + " konnte nicht reingeschrieben werden, weil sie nicht existiert");
            }

        } else {
            System.out.println("Datei existiert bereits!");
            try (BufferedReader bufferedReader = new BufferedReader(new FileReader(PATH_DATEI_STUNDENLOHN))) {
                stundenlohn = bufferedReader.readLine().trim();
                System.out.println("Stundenlon >" + stundenlohn + "< aus Datei gelesen");
            } catch (FileNotFoundException e) {
                System.err.println("Stundenlohn Datei nicht gefunden in Verzeichnis: " + PATH_DATEI_STUNDENLOHN);
            } catch (IOException e) {
                System.out.println("Fehler mit Stundenlohn einlesen aus Datei");
            }
        }

        textFieldStundenlohn.setText(stundenlohn);
    }

    private void saveStundenlohnToDatei(String stundenlohn) {
        Path path = Paths.get(PATH_DATEI_STUNDENLOHN);

        if (!Files.exists(path)) {
            System.out.println("Es gibt die Stundenlohn Datei nicht. Der Wert " + stundenlohn + " kann also nicht gespeichert werden und wird verworfen");
        } else {
            System.out.println("Datei zum speichern des Wertes gefunden!");
            try (PrintWriter printWriter = new PrintWriter(new FileOutputStream(PATH_DATEI_STUNDENLOHN))) {
                printWriter.println(stundenlohn.trim());
                System.out.println("Wert >" + stundenlohn + "< in die existierende Stundenlohn Datei geschrieben");
            } catch (IOException e) {
                System.out.println("In die Stundenlohn Datei im Verzeichnis " + PATH_DATEI_STUNDENLOHN + " konnte nicht reingeschrieben werden (Problem)");
            }
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
                System.out.println("Aud dem Stundenlohn Feld konnte kein gültiger Double entnommen werden. ");
                e.printStackTrace();
                falschesFormatStundenlohn();
                return false;
            } catch (NullPointerException e) {
                System.out.println("Stundenlohn Feld = null");
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
        //textFieldAbrechnungsmonat.setText("");
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
        //textFieldSvBrutto.setText("");
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
        //textFieldMitarbeiternummer.setText("");
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
        //textFieldName.setText("");
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
        lblDateiNichtAkzeptiert.setText("ungültige Eingabe");
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
        lblDateiNichtAkzeptiert.setText("ungültige Eingabe");
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

    protected void fehlgeschlageneSchlussnachricht() {
        lblSchlussnachricht.setText("Der Mitarbeiter hat weniger als 1 Stunde gearbeitet");
        lblSchlussnachricht.setTextFill(Color.RED);
        lblSchlussnachricht.setVisible(true);
    }

    protected void erfolgreicheSchlussnachricht() {
        lblSchlussnachricht.setText("PDF-Dateien wurden erfolgreich generiert");
        lblSchlussnachricht.setTextFill(Color.valueOf("00c300"));
        lblSchlussnachricht.setVisible(true);
    }

    protected void defaultSchlussnachricht() {
        lblSchlussnachricht.setText("");
        lblSchlussnachricht.setTextFill(Color.BLACK);
        lblSchlussnachricht.setVisible(false);
    }


    private static boolean isErrorDisplayed = false;
    public static void displayErrorInGui(String message) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle("Hinweis");
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
        isErrorDisplayed = true;
    }
}
