package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        Parent root = FXMLLoader.load(getClass().getResource("sample.fxml"));
        Scene tapperScene = new Scene(root);
 //      tapperScene.setFill(Color.TRANSPARENT);
//
//
//        primaryStage.initStyle(StageStyle.TRANSPARENT);
        primaryStage.setTitle("Monitor");
        primaryStage.setScene(tapperScene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
