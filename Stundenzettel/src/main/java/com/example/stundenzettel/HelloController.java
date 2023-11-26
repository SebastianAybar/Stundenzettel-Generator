package com.example.stundenzettel;

import javafx.fxml.FXML;
import javafx.scene.control.Label;

public class HelloController {
    @FXML
    private Label label;

    @FXML
    protected void onButtonClickLabelText() {
        label.setText("cweferf");
    }
}