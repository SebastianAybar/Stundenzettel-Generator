module com.example.stundenzettel {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires net.synedra.validatorfx;
    requires de.jensd.fx.glyphs.fontawesome;
    requires org.apache.poi.poi;
    requires org.apache.logging.log4j;
    requires org.apache.pdfbox;
    requires de.focus_shift.jollyday.core;
    requires de.focus_shift.jollyday.jaxb;
    requires org.apache.commons.io;
    requires org.apache.commons.compress;
    requires commons.math3;
    requires itextpdf;
    opens com.example.stundenzettel to javafx.fxml;
    exports com.example.stundenzettel;
}