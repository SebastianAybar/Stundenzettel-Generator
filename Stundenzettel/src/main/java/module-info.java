module com.example.stundenzettel {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires net.synedra.validatorfx;
    requires de.jensd.fx.glyphs.fontawesome;
    requires org.apache.poi.poi;
    //requires eu.hansolo.tilesfx;
    //requires com.almasb.fxgl.all;
    opens com.example.stundenzettel to javafx.fxml;
    exports com.example.stundenzettel;
}