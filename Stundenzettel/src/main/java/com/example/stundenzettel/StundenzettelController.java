package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.HBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
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
            einzelerstellungReader = new EinzelerstellungReader(textFieldAbrechnungsmonat.getText(), textFieldMitarbeiternummer.getText(), textFieldSvBrutto.getText(), textFieldName.getText());
            einzelerstellungReader.writeToExcelEinzelerstellung();
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
    }


}

