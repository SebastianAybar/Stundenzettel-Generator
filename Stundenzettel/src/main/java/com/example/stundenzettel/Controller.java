package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Line;

public class Controller {
    @FXML private AnchorPane anchorPane;
    @FXML private Button excelList;
    @FXML private Button einzelerstellung;
    @FXML private Button calculate;
    @FXML private Button inputPathSearch;
    @FXML private Button outputPathSearch;
    @FXML private Label inputFile;
    @FXML private Label outputTextField;
    @FXML private TextField inputPathTextField;
    @FXML private TextField outputPathTextField;
    @FXML private Line line;
    @FXML private CheckBox replaceFile;
    @FXML private Separator separator1;
    @FXML private Separator separator2;
}