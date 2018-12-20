/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xml.to.xls;

import java.awt.Font;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.image.Image;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javax.swing.filechooser.FileSystemView;

/**
 *
 * @author Arkel
 */
public class XmlToXls extends Application {

    @Override
    public void start(Stage primaryStage) {
        Button btn = new Button();
        btn.setText("Choisir le fichier à convertir (ou glissez et deposez)");
        VBox pane = new VBox();
        primaryStage.getIcons().add(new Image(XmlToXls.class.getResourceAsStream("img/icon.png")));

        try {
//    myDocuments = myDocuments.split("\\s\\s+")[4];
            System.out.println("My documents:" + FileSystemView.getFileSystemView().getDefaultDirectory().getPath());
            File f = new File(FileSystemView.getFileSystemView().getDefaultDirectory().getPath() + "/xml to xls");
            boolean b = f.mkdir();
            System.err.println(b);
        } catch (Exception e) {
            e.printStackTrace();
        }
        ParseConvertUtil pcu = new ParseConvertUtil();
        btn.setOnAction((ActionEvent event) -> {
            System.out.println("Ovrir le fichier à convertir ");
            FileChooser fc = new FileChooser();
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("extensible markup language", "*.xml"));
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("tableurs (xls,xlsx)", "*.xls"));
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("tableurs (xls,xlsx)", "*.xlsx"));

            File f = fc.showOpenDialog(primaryStage);
            if (f != null) {
                String extention = f.getName().substring(f.getName().length() - 3, f.getName().length()).toLowerCase().trim();
                if (extention.equals("xml")) {
                    pcu.readXml(f);
                } else {
                    pcu.readXls(f);
                }
            }
        });

        StackPane root = new StackPane();

        root.getChildren().add(btn);

        String image = XmlToXls.class.getResource("img/background.png").toExternalForm();
        root.setStyle("-fx-background-image: url('" + image + "'); "
                + "-fx-background-position: center center; "
                + "-fx-background-repeat: stretch;");

        Scene scene = new Scene(root, 300, 250);
        scene.setOnDragOver((DragEvent event) -> {
            if (event.getGestureSource() != scene
                    && event.getDragboard().hasFiles()) {
                /* allow for both copying and moving, whatever user chooses */
                event.acceptTransferModes(TransferMode.COPY_OR_MOVE);
            }
            event.consume();
        });
        scene.setOnDragDropped((DragEvent event) -> {
            Dragboard db = event.getDragboard();
            boolean success = false;
            if (db.hasFiles()) {
                db.getFiles().forEach((f) -> {
                    String extention = f.getName().substring(f.getName().length() - 3, f.getName().length()).toLowerCase().trim();
                    String extention2 = f.getName().substring(f.getName().length() - 4, f.getName().length()).toLowerCase().trim();
                    if ("xml".equals(extention)) {
                        pcu.readXml(f);
                    }
                    else if(extention.equals("xls")||extention2.equals("xlsx")){
                        pcu.readXls(f);
                    }
                    else {
                        Alert alert = new Alert(Alert.AlertType.ERROR, "Veuillez déposer des fichiers xml SVP\n" + f.getName() + " n'est pas un fichierr xml ", ButtonType.CLOSE);
                        alert.showAndWait();
                    }
                });
                success = true;
            }
            /* let the source know whether the string was successfully
            * transferred and used */
            event.setDropCompleted(success);

            event.consume();
        });
        primaryStage.setTitle("Xml to Xls");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }

}
