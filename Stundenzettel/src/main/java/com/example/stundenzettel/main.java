package com.example.stundenzettel;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class main extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(main.class.getResource("main.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 525, 400);
        //scene.getStylesheets().add(getClass().getResource("styles.css").toExternalForm());
        String css = this.getClass().getResource("styles.css").toExternalForm();
        scene.getStylesheets().add(css);
        stage.setTitle("Stundenzettel Generator");
        stage.setScene(scene);
        stage.show();
    }

    public static void main(String[] args) {
        launch();
    }
}