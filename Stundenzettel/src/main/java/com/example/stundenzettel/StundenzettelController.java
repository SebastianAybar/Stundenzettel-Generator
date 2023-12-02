package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Line;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.File;
import javafx.scene.image.*;
import java.net.URL;
import java.util.ResourceBundle;
//import javafx.fxml.Initializable;
//import de.jensd.fx.glyphs.fontawesome.FontAwesomeIcon;
//import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;



public class StundenzettelController {
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


//    @Override
//    public void initialize(URL location, ResourceBundle resources) {
//        FontAwesomeIconView iconView = new FontAwesomeIconView(FontAwesomeIcon.STAR);
//        iconView.setSize("12");
//        btnChooseInputFile.setGraphic(iconView);
//    }

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
}