package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.AnchorPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
import java.util.List;

import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;

public class StundenzettelController {
    private AbstractExcelReader abstractExcelReader;
    private AbstractExcelWriter abstractExcelWriter;

    @FXML private AnchorPane anchorPane;
    @FXML private Button excelList;
    @FXML private Button einzelerstellung;

    @FXML private Button btnChooseInputFile;
    @FXML private Button btnChooseOutputFile;
    @FXML private Label inputFile;
    @FXML private Label outputFile;
    @FXML private TextField inputPathTextField;
    @FXML private TextField outputPathTextField;
    @FXML private CheckBox replaceFile;
    @FXML private Separator separator1;
    @FXML private Separator separator2;
    @FXML private Button calculate;
    @FXML private FontAwesomeIconView icnChooseInputFile;
    @FXML private FontAwesomeIconView icnChooseOutputFile;
    @FXML private String inputPath;

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
        abstractExcelReader = new AbstractExcelReader(inputPathTextField.getText());
        List<List<MitarbeiterMonat>> listsOfAbrechnungsmonate = abstractExcelReader.getListsOfAbrechnungsmonate();

        for (List<MitarbeiterMonat> monat : listsOfAbrechnungsmonate) {
            for (MitarbeiterMonat m : monat) {
                System.out.println(m.toString());
            }
            System.out.println();
        }

//        abstractExcelWriter = new AbstractExcelWriter(/*outputPathTextField.getText()*/);
//        abstractExcelWriter.writeToExcel(listsOfAbrechnungsmonate.get(0));
    }




}