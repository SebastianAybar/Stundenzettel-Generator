package com.example.stundenzettel;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;
import java.io.IOException;



public class ApplicationStart extends Application {
    private int windowWidth = 710;
    private int windowHeight = 460;



    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(ApplicationStart.class.getResource("stundenzettel.fxml"));
        Scene scene = new Scene(fxmlLoader.load());

        String css = this.getClass().getResource("styles.css").toExternalForm();
        scene.getStylesheets().add(css);

        stage.setTitle("Stundenzettel Generator");
        stage.setScene(scene);
        stage.show();
        stage.setMinWidth(windowWidth);
        stage.setMinHeight(windowHeight);
    }

    public static void main(String[] args) {
        launch();
    }
}