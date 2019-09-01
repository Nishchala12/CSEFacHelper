package application;

import java.io.IOException;

import javafx.event.ActionEvent;
import javafx.event.Event;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Rectangle2D;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.ImageView;
import javafx.scene.layout.HBox;
import javafx.stage.Screen;
import javafx.stage.Stage;

public class MainController{
	@FXML
	ImageView exit;
	@FXML
	HBox studsetp, markscia, at;
	
	
	
	public void screenDestroy(Event e)
	{
		Stage primstage = (Stage) exit.getScene().getWindow();
		primstage.close();
	}
	
	public void openSetup(Event e)
	{
		System.out.println("das");
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("Controller.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
	
	public void openMarksCIA(Event e)
	{
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("MainFinal.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			primaryStage.setMaximized(true);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
	
	
	public void openAttendance(Event e)
	{
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("Attendance.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setMaximized(true);
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
	
	public void openMarks(Event e)
	{
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("Marks.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			primaryStage.setMaximized(true);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
	
	public void openLabs(Event e)
	{
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("Labs.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			primaryStage.setMaximized(true);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
}