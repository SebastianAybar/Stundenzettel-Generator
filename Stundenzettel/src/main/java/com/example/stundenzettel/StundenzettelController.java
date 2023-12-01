package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Line;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class StundenzettelController {
    @FXML private AnchorPane anchorPane;
    @FXML private Button excelList;
    @FXML private Button einzelerstellung;
    @FXML private Button calculate;
    @FXML private Button btnChooseInputFile;
    @FXML private Button outputPathSearch;
    @FXML private Label inputFile;
    @FXML private Label outputTextField;
    @FXML private TextField inputPathTextField;
    @FXML private TextField outputPathTextField;
    @FXML private Line line;
    @FXML private CheckBox replaceFile;
    @FXML private Separator separator1;
    @FXML private Separator separator2;

    @FXML
    protected void chooseFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Datei auswählen");
        File selectedFile = fileChooser.showOpenDialog(null);

        if (selectedFile != null) {
            inputPathTextField.setText(selectedFile.getAbsolutePath());
        }
    }
}