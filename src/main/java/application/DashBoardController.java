package application;

import java.io.BufferedReader;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Random;
import java.util.concurrent.CountDownLatch;
import java.util.stream.Collectors;

import javafx.application.Platform;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.Event;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.geometry.Rectangle2D;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ButtonType;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.ListView.EditEvent;
import javafx.scene.control.TableColumn.CellEditEvent;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldListCell;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.util.Pair;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.List;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import com.google.auth.oauth2.GoogleCredentials;
import com.google.firebase.FirebaseApp;
import com.google.firebase.FirebaseOptions;
import com.google.firebase.database.DataSnapshot;
import com.google.firebase.database.DatabaseError;
import com.google.firebase.database.DatabaseReference;
import com.google.firebase.database.FirebaseDatabase;
import com.google.firebase.database.ValueEventListener;

import application.Controller.Person;
public class DashBoardController {
	
	@FXML 
	AnchorPane ap, ap_attendance, header, ap_lab,ap_marks;
	
	
	@FXML
	ImageView i1,i2,i3,i4;
	
	String rootpath = "C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM";
	@FXML
	VBox utility_lab, utility_attend, utility_marks;
	@FXML
	HBox  qa,at,mark_nav,studset;
	
	String coursecode ="";
	//Attendance
	@FXML
	private TableView<Person> table = new TableView<Person>();
	
	@FXML
	private TableView<Person_Marks> table_marks;
	

	
	 private ObservableList<Person_Marks> data1 =
		        FXCollections.observableArrayList();
	
TableColumn usnCol1,nameCol1,marksCol,newCol;

	@FXML
	DatePicker datePicker;
	
	   @FXML
	   private TextField addTotalClasses;
	    
	    private ObservableList<Person> data =
		        FXCollections.observableArrayList();
	    TableColumn usnCol,nameCol,classesCol,perCol,usnCol2,nameCol2;
	    ArrayList<String> ref=new ArrayList<>();
	    
	    @FXML
	    Button save,load;
	    @FXML
	    AnchorPane ap_calendar;
	//-----------------------------------------    
	    
	@FXML
	Button savespbtn,loadspbtn, savefir, syncsave, loadlab, savelab, importMoodle,saveFirLab, loadFirLab, savefiremarksbtn, loadfiremarksbtn, loadmarksbtn, savemarksbtn;
	
	@FXML
	Button btn;
	ListView lab1 = new ListView<String>();
    ListView lab2 = new ListView<String>();
    ListView lab3 = new ListView<String>();
    ListView lab4 = new ListView<String>();
    ListView lab5 = new ListView<String>();
    ListView lab6 = new ListView<String>();
    ListView lab7 = new ListView<String>();
    ListView lab8 = new ListView<String>();
    ListView lab9 = new ListView<String>();
    ListView lab10 = new ListView<String>();
    ListView listView0 = new ListView<String>();
    ListView listView1 = new ListView<String>();
    ListView listView2 = new ListView<String>();
    ListView listView3 = new ListView<String>();
    
    
   
    
    ArrayList<ArrayList<String>> marks;
	ArrayList<String> names;
	
	Stage primaryStage = new Stage();
	HBox root = new HBox();

    // Declare ListView's
	
    
    VBox vsl = new VBox();
    VBox vusn = new VBox();
    VBox vname = new VBox();
    VBox vcat = new VBox();
    
    VBox vp1 = new VBox();
    VBox vp2 = new VBox();
    VBox vp3 = new VBox();
    VBox vp4 = new VBox();
    VBox vp5 = new VBox();
    VBox vp6 = new VBox();
    VBox vp7 = new VBox();
    VBox vp8 = new VBox();
    VBox vp9 = new VBox();
    VBox vp10 = new VBox();
    
   
    
    Label lsl,lsusn,lname,lcat,lp1,lp2,lp3,lp4,lp5,lp6,lp7,lp8,lp9,lp10;
    
    ObservableList<String> lap1 = FXCollections.observableArrayList();
    ObservableList<String> lap2 = FXCollections.observableArrayList();
    ObservableList<String> lap4 = FXCollections.observableArrayList();
    ObservableList<String> lap3 = FXCollections.observableArrayList();
    ObservableList<String> lap5 = FXCollections.observableArrayList();
    ObservableList<String> lap6 = FXCollections.observableArrayList();
    ObservableList<String> lap7 = FXCollections.observableArrayList();
    ObservableList<String> lap8 = FXCollections.observableArrayList();
    ObservableList<String> lap9 = FXCollections.observableArrayList();
    ObservableList<String> lap10 = FXCollections.observableArrayList();
    
    Rectangle2D screenBounds = Screen.getPrimary().getBounds();
    ArrayList<String> usns = new ArrayList<>();
    ArrayList<String> name = new ArrayList<>();
    static ArrayList<String> studdat = new ArrayList<String>();
    
	public void initialize() throws IOException
	{		
		
		
		 Rectangle2D screenBounds = Screen.getPrimary().getBounds();
	       //header.setPrefWidth(screenBounds.getWidth());
	       utility_lab.setPrefHeight(screenBounds.getHeight());
	      // vbox_nav.setPrefHeight(screenBounds.getHeight());
	        
	        
	        int width = (int) screenBounds.getWidth();
	        int utilsize = (width/100)*20;
	        int lab_size = (width/100)*60;
	        utility_lab.setPrefWidth(utilsize);
	        utility_attend.setPrefWidth(utilsize);
	        utility_marks.setPrefWidth(utilsize);
	        //vbox_nav.setPrefWidth(width-utilsize-lab_size);
	        table.setPrefWidth(lab_size);
	        table.setPrefHeight(screenBounds.getHeight());
	        
	        table_marks.setPrefWidth(lab_size);
	        table_marks.setPrefHeight(screenBounds.getHeight());
	        
	        utility_attend.setPrefHeight(screenBounds.getHeight());
	        utility_marks.setPrefHeight(screenBounds.getHeight());
	        ap.setPrefWidth(lab_size);
	        ap_marks.setPrefWidth(lab_size);
	        
	        ap_marks.setPrefHeight(screenBounds.getHeight());
	        System.out.println(utilsize);
	        
	        
	        //at.setPrefWidth(width-utilsize-lab_size);
	        //qa.setPrefWidth(width-utilsize-lab_size);
	        //mark_nav.setPrefWidth(width-utilsize-lab_size);
	        //studset.setPrefWidth(width-utilsize-lab_size);
	        
	        
	        savespbtn.setPrefWidth(utilsize);
	        loadspbtn.setPrefWidth(utilsize);
	        savefir.setPrefWidth(utilsize);
	        syncsave.setPrefWidth(utilsize);
	        addTotalClasses.setPrefWidth(utilsize);
	        ap_calendar.setPrefWidth(utilsize);

	        loadlab.setPrefWidth(utilsize);
	        savelab.setPrefWidth(utilsize);
	        importMoodle.setPrefWidth(utilsize);
	        loadFirLab.setPrefWidth(utilsize);
	        saveFirLab.setPrefWidth(utilsize);
	        
	        loadfiremarksbtn.setPrefWidth(utilsize);
	        savefiremarksbtn.setPrefWidth(utilsize);
	        savemarksbtn.setPrefWidth(utilsize);
	        loadmarksbtn.setPrefWidth(utilsize);
	        
	        table.setEditable(true);
	    	
			usnCol1 = new TableColumn("USN");
	        usnCol1.setMinWidth(100);
	        usnCol1.setCellValueFactory(
	                new PropertyValueFactory<Person, String>("usn"));
	       
	        nameCol1 = new TableColumn("NAME");
	        nameCol1.setMinWidth(100);
	        nameCol1.setCellValueFactory(
	                new PropertyValueFactory<Person, String>("name"));
	 
	        classesCol = new TableColumn("Classes Attended");
	        classesCol.setMinWidth(150);
	        classesCol.setCellValueFactory(
	                new PropertyValueFactory<Person, String>("classes"));
	        classesCol.setCellFactory(TextFieldTableCell.forTableColumn());
	        classesCol.setOnEditCommit(
	        new EventHandler<CellEditEvent<Person, String>>() {
	        @Override
	        public void handle(CellEditEvent<Person, String> t) {
	        ((Person) t.getTableView().getItems().get(
	        t.getTablePosition().getRow())
	        ).setClasses(t.getNewValue());
	        }
	        }
	        );
	        
	        perCol = new TableColumn("Percentage");
	        perCol.setMinWidth(150);
	        perCol.setCellValueFactory(
	                new PropertyValueFactory<Person, String>("PERCENTAGE"));
	        
	        table.setItems(data);
	        table.getColumns().addAll(usnCol1,nameCol1, classesCol);
	        
	        
	        table_marks.setEditable(true);
	        usnCol2 = new TableColumn("USN");
	        usnCol2.setMinWidth(100);
	        usnCol2.setCellValueFactory(
	                new PropertyValueFactory<Person_Marks, String>("usn"));
	       
	        nameCol2 = new TableColumn("NAME");
	        nameCol2.setMinWidth(100);
	        nameCol2.setCellValueFactory(
	                new PropertyValueFactory<Person_Marks, String>("name"));

	        marksCol = new TableColumn("Marks[50]");
	        marksCol.setMinWidth(100);
	        marksCol.setCellValueFactory(
	                new PropertyValueFactory<Person_Marks, String>("marks"));
	        marksCol.setCellFactory(TextFieldTableCell.forTableColumn());
	        marksCol.setOnEditCommit(
	        new EventHandler<CellEditEvent<Person_Marks, String>>() {
	        @Override
	        public void handle(CellEditEvent<Person_Marks, String> t) {
	        ((Person_Marks) t.getTableView().getItems().get(
	        t.getTablePosition().getRow())
	        ).setMarks(t.getNewValue());
	        }
	        }
	        );
	        
	        table_marks.setItems(data1);
	        table_marks.getColumns().addAll(usnCol2,nameCol2, marksCol);
	        
	        
	        Label lbl = new Label("dd/mm/yyyy");
	    	datePicker = new DatePicker();

	    	datePicker.setOnAction(e -> {
	    	LocalDate date = datePicker.getValue();
	    	DateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
	    	Date conv_date = java.sql.Date.valueOf(date);
	    	String finalDate = formatter.format(conv_date);
	    	System.out.println(finalDate);
	    	finalDate = finalDate.replace('/', '-');
	    	lbl.setText(finalDate);
	    	});
	    	
	    	ap_calendar.getChildren().add(datePicker);
	        
	   
	        
	}

	public void marksShow(ActionEvent e) throws IOException
	{
		ap_attendance.setVisible(false);	
		ap_lab.setVisible(false);
		ap_marks.setVisible(true);
	}
	
	
	
	public void newwin(ActionEvent e) throws IOException
	{
		ap_attendance.setVisible(false);	
		ap_lab.setVisible(true);	
		ap_marks.setVisible(false);
	}
	
	public void attendanceShow(ActionEvent e) throws IOException
	{
		ap_lab.setVisible(false);
		ap_marks.setVisible(false);
		ap_attendance.setVisible(true);	
	}
	
	
	public  void LoadMarks(ActionEvent e)throws IOException
	{
		Dialog<Pair<String, String>> dialog = new Dialog<>();
	    dialog.setTitle("TestName");

	    // Set the button types.
	    ButtonType loginButtonType = new ButtonType("OK", ButtonData.OK_DONE);
	    dialog.getDialogPane().getButtonTypes().addAll(loginButtonType, ButtonType.CANCEL);

	            GridPane gridPane = new GridPane();
	    gridPane.setHgap(10);
	    gridPane.setVgap(10);
	    gridPane.setPadding(new Insets(20, 150, 10, 10));

	    TextField from = new TextField();
	    from.setPromptText("From");
	    TextField to = new TextField();
	    to.setPromptText("To");

	    gridPane.add(new Label("Semester:"), 0, 0);
	    gridPane.add(from, 1, 0);
	    gridPane.add(new Label("Section:"), 2, 0);
	    gridPane.add(to, 3, 0);

	    dialog.getDialogPane().setContent(gridPane);

	    // Request focus on the username field by default.
	    Platform.runLater(() -> from.requestFocus());

	    // Convert the result to a username-password-pair when the login button is clicked.
	    dialog.setResultConverter(dialogButton -> {
	        if (dialogButton == loginButtonType) {
	            return new Pair<>(from.getText(), to.getText());
	        }
	        return null;
	    });

	    Optional<Pair<String, String>> result = dialog.showAndWait();
	    
	    studdat.clear();
	    
	    result.ifPresent(pair -> {
	    	String sem = pair.getKey().toUpperCase();
	    	String sec = pair.getValue().toUpperCase();
	    	studdat.add(sem);
	    	studdat.add(sec);
	        System.out.println("From=" + pair.getKey() + ", To=" + pair.getValue());
	    });
		
	    
	    table_marks.getColumns().clear();
	    table_marks.getColumns().addAll(usnCol2, nameCol2, marksCol);
	    table_marks.setItems(data1);
		
		InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		HSSFSheet sheet = wb.getSheetAt(0);
		HSSFRow row; 
		data1.clear();
		for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
		{
		
		HSSFRow row1 = sheet.getRow(i);
		Iterator cells = row1.cellIterator();
		while (cells.hasNext())
		{
		
		cells.next();
		}
		String usn = sheet.getRow(i).getCell(0).toString();
		String name = sheet.getRow(i).getCell(1).toString();
		String marks50 = sheet.getRow(i).getCell(2).toString();
		
		
		
		data1.add(new Person_Marks(usn,name,marks50));
		
			
	}
		
		table_marks.getColumns().clear();
		table_marks.getColumns().addAll(usnCol2, nameCol2, marksCol);
		table_marks.setItems(data1);
		
	}
	
	public  void SaveMarks(ActionEvent e)throws IOException
	{
		InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook  workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet spreadsheet = workbook.getSheetAt(0);


	    
        int i=1;
        for(Person_Marks dsce: data1)
		{
			if(!dsce.getMarks().equals(""))
			{
				spreadsheet.getRow(i).createCell(2).setCellValue(dsce.getMarks());
			System.out.println(dsce.getMarks());
			int mark = Integer.parseInt((dsce.getMarks().toString()));
			
			
			int marks = (int) Math.round(mark*0.2);;
			
			spreadsheet.getRow(i).createCell(3).setCellValue(marks+"");
			}
			i++;
			
		}

	    

	    FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+studdat.get(0)+studdat.get(1)+".xls");
	    workbook.write(fileOut);
	    fileOut.close();
	    
	    Alert alert=new Alert(AlertType.INFORMATION);
	    alert.setTitle("Information Dialog");
	    alert.setHeaderText(null);
	    alert.setContentText("Data Saved!");
	    alert.showAndWait();
		return;
	}
	
	public void SaveFirebaseAttendance(ActionEvent e) throws IOException
	{
		
		
		
		
		TextInputDialog dialog = new TextInputDialog("Tran");
		 
		dialog.setTitle("Save to Firebase");
		dialog.setHeaderText("Enter the course code:");
		dialog.setContentText("Name:");
		 
		Optional<String> result = dialog.showAndWait();
		 
		result.ifPresent(name -> {
			coursecode = name;
		});
	    
		int tc = Integer.parseInt(addTotalClasses.getText().toString());

   	 ArrayList<String> attend = new ArrayList<>();
   	 ArrayList<String> percent = new ArrayList<>();
		
		for(Person dsce: data)
		{
			if(!dsce.getClasses().equals(""))
			{
			attend.add(dsce.getClasses().toString());
			Double percentage = Double.parseDouble((dsce.getClasses().toString()))/tc;
			percentage= percentage *100;
			
			int perc = (int) Math.round(percentage);
			
			 percent.add(perc+"");
			}
			
			
		}
		
		
	    try {
	    	
	    	ArrayList<ArrayList<String>> big = new ArrayList<ArrayList<String>>();
	      	
	      	
	      	
	      	
	            final CountDownLatch latch1 = new CountDownLatch(1);
	            DatabaseReference ref1= FirebaseDatabase.getInstance().getReference();
	            DatabaseReference ref2;    
	             ref2 = ref1.child("MARKS");

	        	 String tchr_name = coursecode;
	        	
	        	
	        	 String att = String.join(",", attend);
	        	 String perc = String.join(",", percent);
	        	 DatabaseReference ref = FirebaseDatabase.getInstance().getReference("Marks/sem_2/D/"+tchr_name);
	        	 	 DatabaseReference child_name = FirebaseDatabase.getInstance().getReference();
	        	 	
	        	 child_name=ref.child("att");
	        	 child_name.setValueAsync(att);
	        	 child_name=ref.child("perc");
	        	 child_name.setValueAsync(perc);
	        	 latch1.countDown();
	        	 
	        	System.out.println("Succesfull");
	        	 
	        	latch1.await();
	    			   } 
	    			 catch (InterruptedException ef) {
	    			        ef.printStackTrace();
	    			    }
			
	}
	
	
	
	
	public void LoadFirebaseAttendance(ActionEvent e) throws IOException
	{

	     ArrayList<DataSnapshot> Userlist = new ArrayList<DataSnapshot>();   
	      try {
	              final CountDownLatch latch1 = new CountDownLatch(1);
	              DatabaseReference ref1= FirebaseDatabase.getInstance().getReference();
	              DatabaseReference ref2;    
	               ref2 = ref1.child("Marks/sem_2/");


	               ref2.addListenerForSingleValueEvent(
	            		   new ValueEventListener() {
	                public void onDataChange(DataSnapshot dataSnapshot) {

	                    //ArrayList<Object> Userlist = new ArrayList<Object>();   
	                    ArrayList<ArrayList<String>> big_arr = new ArrayList<ArrayList<String>>();
	                   	                      for (DataSnapshot dsp : dataSnapshot.getChildren()) {
	                          Userlist.add(dsp); 
	                         
	                        }
	                    //big_arr.add(Userlist);
	                    
	         				 // System.out.println(Userlist.get(0)+"dsad"+Userlist.size());
	         				     
	                                                latch1.countDown();
	     				     }

	    			        	  public void onCancelled(DatabaseError error) {
	    			        		  latch1.countDown();
	    			        		  
	    			        	  }
	    			        	});
	    			        	 latch1.await();
	     			   } 
	     			 catch (InterruptedException en) {
	    				        en.printStackTrace();
	    				    }
	     			
	      	ArrayList<ArrayList<String>> big = new ArrayList<ArrayList<String>>();
	      	ArrayList<String> smol = new ArrayList<String>();
	      	
	      	for(DataSnapshot d: Userlist.get(0).getChildren())
	      	{
	      		smol = new ArrayList<String>();
	      		FireData fir = d.getValue(FireData.class);
	      		System.out.println(fir.getAtt());
	      		smol.addAll(Arrays.asList(fir.getAtt().split(",")));
	      		big.add(smol);
	      		smol = new ArrayList<String>();
	      		smol.addAll(Arrays.asList(fir.getPerc().split(",")));
	      		big.add(smol);
	      	}
	    	
	      	System.out.println(big);
	      	combine(big);
	
	}	
	
	public void saveAttendance(ActionEvent e) throws IOException
	{
	 
    	 String a=addTotalClasses.getText();
    	 
 	    int tc=Integer.parseInt(a);
 	   
 	    
 	   InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook  workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet spreadsheet = workbook.getSheetAt(0);

        

        System.out.println(table.getItems().size());
        
        int i=1;
        for(Person dsce: data)
		{
			if(!dsce.getClasses().equals(""))
			{
				spreadsheet.getRow(i).createCell(2).setCellValue(dsce.getClasses());
			System.out.println(dsce.getClasses());
			Double percentage = Double.parseDouble((dsce.getClasses().toString()))/tc;
			percentage= percentage *100;
			
			int perc = (int) Math.round(percentage);
			
			spreadsheet.getRow(i).createCell(3).setCellValue(perc+"");
			}
			i++;
			
		}
        
       
       
        LocalDate date=java.time.LocalDate.now();
        DateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        Date conv_date = java.sql.Date.valueOf(date);
        String finalDate = formatter.format(conv_date);
        finalDate = finalDate.replace('/', '-');
  
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Nisha\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+".xls");
        workbook.write(fileOut);
        fileOut.close();
        

        Alert alert=new Alert(AlertType.INFORMATION);
        alert.setTitle("Information Dialog");
        alert.setHeaderText(null);
        alert.setContentText("Spreadsheet Saved!");
        alert.showAndWait();
        }

public void loadAttendance(ActionEvent e)throws IOException
{

	Dialog<Pair<String, String>> dialog = new Dialog<>();
    dialog.setTitle("TestName");

    // Set the button types.
    ButtonType loginButtonType = new ButtonType("OK", ButtonData.OK_DONE);
    dialog.getDialogPane().getButtonTypes().addAll(loginButtonType, ButtonType.CANCEL);

            GridPane gridPane = new GridPane();
    gridPane.setHgap(10);
    gridPane.setVgap(10);
    gridPane.setPadding(new Insets(20, 150, 10, 10));

    TextField from = new TextField();
    from.setPromptText("From");
    TextField to = new TextField();
    to.setPromptText("To");

    gridPane.add(new Label("Semester:"), 0, 0);
    gridPane.add(from, 1, 0);
    gridPane.add(new Label("Section:"), 2, 0);
    gridPane.add(to, 3, 0);

    dialog.getDialogPane().setContent(gridPane);

    // Request focus on the username field by default.
    Platform.runLater(() -> from.requestFocus());

    // Convert the result to a username-password-pair when the login button is clicked.
    dialog.setResultConverter(dialogButton -> {
        if (dialogButton == loginButtonType) {
            return new Pair<>(from.getText(), to.getText());
        }
        return null;
    });

    Optional<Pair<String, String>> result = dialog.showAndWait();
    
    studdat.clear();
    
    result.ifPresent(pair -> {
    	String sem = pair.getKey().toUpperCase();
    	String sec = pair.getValue().toUpperCase();
    	studdat.add(sem);
    	studdat.add(sec);
        System.out.println("From=" + pair.getKey() + ", To=" + pair.getValue());
    });
	
    
	table.getColumns().clear();
	table.getColumns().addAll(usnCol1, nameCol1, classesCol);
	table.setItems(data);
	
	String finalDate0="";
	
     LocalDate date = datePicker.getValue();
     DateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
     Date conv_date = java.sql.Date.valueOf(date);
     finalDate0 = formatter.format(conv_date);
     finalDate0 = finalDate0.replace('/', '-');
	
	data.clear();
	table.setItems(data);
	
	String[] sheetrows ;
		
		InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook  wb = new HSSFWorkbook(ExcelFileToRead);
		HSSFSheet sheet = wb.getSheetAt(0);
		
		
		HSSFRow row; 
		HSSFCell cell;

		Iterator rows = sheet.rowIterator();
		
		sheetrows = new String[3];
		
		int k =1;
		System.out.println(sheet.getPhysicalNumberOfRows());
		while(k<sheet.getPhysicalNumberOfRows())
		{
			data.add(new Person(sheet.getRow(k).getCell(0).getStringCellValue(),
					sheet.getRow(k).getCell(1).getStringCellValue(),
					sheet.getRow(k).getCell(2).getStringCellValue()));
			//System.out.println(sheet.getRow(k).getCell(2).getStringCellValue());
			k++;
		}
		table.setItems(data);
}

	
   
	
	
	public void saveList(ActionEvent e) throws NumberFormatException, IOException
	{
		
		
		int n;
		XWPFDocument docX2 = new XWPFDocument();
	
		   BufferedReader in= new BufferedReader(new InputStreamReader(System.in));
	    
	      FileOutputStream out = new FileOutputStream(new File("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+studdat.get(0)+studdat.get(1)+".docx"));
	      
	      XWPFParagraph paragraph = docX2.createParagraph();
	      paragraph.setAlignment(ParagraphAlignment.CENTER);
	      
	      XWPFRun paragraphOneRunOne = paragraph.createRun();
	      paragraphOneRunOne.setBold(true);
	      paragraphOneRunOne.setText("DAYANANDA SAGAR COLLEGE OF ENGINEERING");
	      paragraphOneRunOne.addBreak();
	      
	     
	      XWPFRun paragraphOneRunTwo = paragraph.createRun();
	      paragraphOneRunTwo.setItalic(true);
	      paragraphOneRunTwo.setText("(An autonomous Institute to VTU,Belagavi)");
	      paragraphOneRunTwo.addBreak();
	      
	      XWPFRun paragraphOneRunThree = paragraph.createRun();
	      paragraphOneRunThree.setBold(false);
	      paragraphOneRunThree.setText("Shavige Malleshwara Hills,Kumarswamy Layout,Bengaluru-560078");
	      paragraphOneRunThree.addBreak();
	      
	      XWPFRun paragraphOneRunFour = paragraph.createRun();
	      paragraphOneRunFour.setBold(true);
	      paragraphOneRunFour.setText("Academic Session: Jan 2019-May 2019");
	      paragraphOneRunFour.addBreak();
	      paragraphOneRunFour.addBreak();
	      paragraphOneRunFour.addBreak();
	      
	      XWPFRun paragraphOneRunFive= paragraph.createRun();
	      paragraphOneRunFive.setBold(true);
	      paragraphOneRunFive.setText("Laboratory");
	      
	    
	      
	
	n= usns.size();
	ChangeOrientation obj= new ChangeOrientation();

    
	/*switch(p)
	{ 
	case 1:  obj.edit_t (n,docX2,4,10.0,10.0 ,10.0, 10.0);
	         break;
	case 2:  obj.edit_t(n,docX2,5,10.0,10.0 ,10.0, 10.0);
	         break;
	case 3:  obj.edit_t (n,docX2,6 ,10.0,10.0 ,10.0, 10.0);
	         break; 
	case 4:  obj.edit_t (n,docX2,7 ,10.0,10.0 ,10.0, 10.0);
	         break;
	case 5:  obj.edit_t (n,docX2,8 ,10.0,10.0 ,10.0, 10.0);
	         break;
	case 6:  obj.edit_t (n,docX2,9 ,10.0,10.0 ,10.0, 10.0);
	         break;
	case 7:  obj.edit_t (n,docX2,10,10.0,10.0 ,10.0, 10.0);
	         break;
	case 8:  obj.edit_t (n,docX2,11 ,10.0,10.0 ,10.0, 10.0);
	         break;
	case 9:  obj.edit_t (n,docX2,12 ,10.0,10.0 ,10.0, 10.0);
	         break;
	case 10:  obj.edit_t (n,docX2,13 ,10.0,10.0 ,10.0, 10.0);
	         break; 
	case 11:  obj.edit_t (n,docX2,14,10.0,10.0 ,10.0, 10.0);
	         break; 
	case 12:  obj.edit_t (n,docX2,15,10.0,10.0 ,10.0, 10.0);
	         break;
	}

	*/
	List<String> usn = listView1.getItems();
	List<String> names = listView2.getItems();
	List<String> p1 = lab1.getItems();
	List<String> p2 = lab2.getItems();
	List<String> p3 = lab3.getItems();
	List<String> p4 = lab4.getItems();
	List<String> p5 = lab5.getItems();
	List<String> p6 = lab6.getItems();
	List<String> p7 = lab7.getItems();
	List<String> p8 = lab8.getItems();
	List<String> p9 = lab9.getItems();
	List<String> p10 = lab10.getItems();
	
	
	obj.populateDoc(docX2, usn, names, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10);
	
	
	docX2.write(out);
	            out.close();
	            
	            System.out.println(".docx written successully");
	         
	         
	             
	
	saveSpreadSheet();
		
		
		
	}
	
	public void saveSpreadSheet() throws IOException {
		InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet sheet = workbook.getSheetAt(0);
        
        List<String> usn = listView1.getItems();
    	List<String> names = listView2.getItems();
    	List<String> p1 = lab1.getItems();
    	List<String> p2 = lab2.getItems();
    	List<String> p3 = lab3.getItems();
    	List<String> p4 = lab4.getItems();
    	List<String> p5 = lab5.getItems();
    	List<String> p6 = lab6.getItems();
    	List<String> p7 = lab7.getItems();
    	List<String> p8 = lab8.getItems();
    	List<String> p9 = lab9.getItems();
    	List<String> p10 = lab10.getItems();
    	
    	
        for(int i=0;i<p1.size();i++)
        {
        	sheet.createRow(i).createCell(0).setCellValue(p1.get(i)+"");
        	sheet.getRow(i).createCell(1).setCellValue(p2.get(i)+"");
        	sheet.getRow(i).createCell(2).setCellValue(p3.get(i)+"");
        	sheet.getRow(i).createCell(3).setCellValue(p4.get(i)+"");
        	sheet.getRow(i).createCell(4).setCellValue(p5.get(i)+"");
        	sheet.getRow(i).createCell(5).setCellValue(p6.get(i)+"");
        	sheet.getRow(i).createCell(6).setCellValue(p7.get(i)+"");
        	sheet.getRow(i).createCell(7).setCellValue(p8.get(i)+"");
        	sheet.getRow(i).createCell(8).setCellValue(p9.get(i)+"");
        	sheet.getRow(i).createCell(9).setCellValue(p10.get(i)+"");
        	
        }
        
        
        try {
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+studdat.get(0)+studdat.get(1)+".xls");
        workbook.write(fileOut);
        fileOut.close();
        }
        catch(Exception er)
        {
        	
        }
        
    	Alert alert = new Alert(AlertType.INFORMATION, "Saved the Model File to Project Location", ButtonType.OK);
        alert.getDialogPane().setMinHeight(Region.USE_PREF_SIZE);
        alert.show();	
	}
	
	public void loadSpreadSheet(ActionEvent e) throws IOException
	{
		
		Dialog<Pair<String, String>> dialog = new Dialog<>();
	    dialog.setTitle("TestName");

	    // Set the button types.
	    ButtonType loginButtonType = new ButtonType("OK", ButtonData.OK_DONE);
	    dialog.getDialogPane().getButtonTypes().addAll(loginButtonType, ButtonType.CANCEL);

	            GridPane gridPane = new GridPane();
	    gridPane.setHgap(10);
	    gridPane.setVgap(10);
	    gridPane.setPadding(new Insets(20, 150, 10, 10));

	    TextField from = new TextField();
	    from.setPromptText("From");
	    TextField to = new TextField();
	    to.setPromptText("To");

	    gridPane.add(new Label("Semester:"), 0, 0);
	    gridPane.add(from, 1, 0);
	    gridPane.add(new Label("Section:"), 2, 0);
	    gridPane.add(to, 3, 0);

	    dialog.getDialogPane().setContent(gridPane);

	    // Request focus on the username field by default.
	    Platform.runLater(() -> from.requestFocus());

	    // Convert the result to a username-password-pair when the login button is clicked.
	    dialog.setResultConverter(dialogButton -> {
	        if (dialogButton == loginButtonType) {
	            return new Pair<>(from.getText(), to.getText());
	        }
	        return null;
	    });

	    Optional<Pair<String, String>> result = dialog.showAndWait();
	    
	    
	    
	    result.ifPresent(pair -> {
	    	String sem = pair.getKey().toUpperCase();
	    	String sec = pair.getValue().toUpperCase();
	    	studdat.add(sem);
	    	studdat.add(sec);
	        System.out.println("From=" + pair.getKey() + ", To=" + pair.getValue());
	    });
	    
        listView0.setEditable(true);
        listView0.setCellFactory(TextFieldListCell.forListView());		
        listView0.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				listView0.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        listView0.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {

			@Override
			public void handle(ListView.EditEvent<String> t) {
				System.out.println("setOnEditCancel");
			}
		});
        
        listView1.setEditable(true);
        listView1.setCellFactory(TextFieldListCell.forListView());		
        listView1.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				listView1.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        listView1.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {

			@Override
			public void handle(ListView.EditEvent<String> t) {
				System.out.println("setOnEditCancel");
			}
		});
        
        listView2.setEditable(true);
        listView2.setCellFactory(TextFieldListCell.forListView());		
        listView2.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				listView2.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        listView2.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {

			@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        listView3.setEditable(true);
        listView3.setCellFactory(TextFieldListCell.forListView());		
        listView3.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				listView3.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        listView3.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
       
        generateListEditable();
        
        
        
        
        BufferedReader br1 = new BufferedReader(new FileReader(rootpath+"\\StudentData\\"+studdat.get(0)+studdat.get(1)+".txt"));
		 int m =1;
		 String line;
		while((line = br1.readLine()) !=null)
	    	{
	    		String eachLine = line;
	    		usns.add(m+"");
	    		name.add(line);
	    		m++;
	    	}
        
        // Fill Data inside ListView
        ArrayList<String> dataarr = new ArrayList<>();
        ArrayList<String> dataarr1 = new ArrayList<>();
		for(int i=0;i<10;i++)
		{
			dataarr.add(i+"");
		}
        ObservableList<String> data = FXCollections.observableArrayList(usns);
        ObservableList<String> names = FXCollections.observableArrayList(name);
        
        
        Random rand = new Random();
        
        for(int i=0;i<usns.size();i++)
		{
			dataarr1.add("C");
			dataarr1.add("V");
			dataarr1.add("R");
			dataarr1.add("T");
			lap1.add("");
			lap1.add("");
			lap1.add("");
			lap1.add("");
			
			lap2.add("");
			lap2.add("");
			lap2.add("");
			lap2.add("");
			
			lap3.add("");
			lap3.add("");
			lap3.add("");
			lap3.add("");
			
			lap4.add("");
			lap4.add("");
			lap4.add("");
			lap4.add("");
			
			lap5.add("");
			lap5.add("");
			lap5.add("");
			lap5.add("");
			
			lap6.add("");
			lap6.add("");
			lap6.add("");
			lap6.add("");
			
			lap7.add("");
			lap7.add("");
			lap7.add("");
			lap7.add("");
			
			lap8.add("");
			lap8.add("");
			lap8.add("");
			lap8.add("");
			
			lap9.add("");
			lap9.add("");
			lap9.add("");
			lap9.add("");
			
			lap10.add("");
			lap10.add("");
			lap10.add("");
			lap10.add("");
		}
        ObservableList<String> dat1 = FXCollections.observableArrayList(dataarr1);
		
        
        listView0.setItems(data);
        listView1.setItems(data);
        listView2.setItems(names);
        listView3.setItems(dat1);
        

        lab1.setItems(lap1);
        lab2.setItems(lap2);
        lab3.setItems(lap3);
        lab4.setItems(lap4);
        lab5.setItems(lap5);
        lab6.setItems(lap6);
        lab7.setItems(lap7);
        lab8.setItems(lap8);
        lab9.setItems(lap9);
        lab10.setItems(lap10);
        
        listView0.setFixedCellSize(96);
        listView1.setFixedCellSize(96);
        listView2.setFixedCellSize(96);
        listView3.setFixedCellSize(24);
        
        
        lsl=new Label("Sl #");
        lsusn=new Label("USN");
        lname=new Label("Name");
        lcat= new Label("");
        lp1=new Label("P1");
        lp2=new Label("P2");
        lp3=new Label("P3");
        lp4=new Label("P4");
        lp5=new Label("P5");
        lp6=new Label("P6");
        lp7=new Label("P7");
        lp8=new Label("P8");
        lp9=new Label("P9");
        lp10=new Label("P10");
        
        
        vsl.getChildren().addAll(lsl,listView0);
        vusn.getChildren().addAll(lsusn,listView1);
        vname.getChildren().addAll(lname,listView2);
        vcat.getChildren().addAll(lcat,listView3);
        vp1.getChildren().addAll(lp1, lab1);
        vp2.getChildren().addAll(lp2, lab2);
        vp3.getChildren().addAll(lp3, lab3);
        vp4.getChildren().addAll(lp4, lab4);
        vp5.getChildren().addAll(lp5, lab5);
        vp6.getChildren().addAll(lp6, lab6);
        vp7.getChildren().addAll(lp7, lab7);
        vp8.getChildren().addAll(lp8, lab8);
        vp9.getChildren().addAll(lp9, lab9);
        vp10.getChildren().addAll(lp10, lab10);
        
        Rectangle2D screenBounds = Screen.getPrimary().getBounds();
        listView0.setPrefHeight(screenBounds.getHeight());
        listView1.setPrefHeight(screenBounds.getHeight());
        listView2.setPrefHeight(screenBounds.getHeight());
        listView3.setPrefHeight(screenBounds.getHeight());
        
        lab1.setPrefHeight(screenBounds.getHeight());
        lab2.setPrefHeight(screenBounds.getHeight());
        lab3.setPrefHeight(screenBounds.getHeight());
        lab4.setPrefHeight(screenBounds.getHeight());
        lab5.setPrefHeight(screenBounds.getHeight());
        lab6.setPrefHeight(screenBounds.getHeight());
        lab7.setPrefHeight(screenBounds.getHeight());
        lab8.setPrefHeight(screenBounds.getHeight());
        lab9.setPrefHeight(screenBounds.getHeight());
        lab10.setPrefHeight(screenBounds.getHeight());
        
        listView0.setPrefWidth(80);
        listView3.setPrefWidth(100);
        
        lab1.setPrefWidth(100);
        lab2.setPrefWidth(100);
        lab3.setPrefWidth(100);
        lab4.setPrefWidth(100);
        lab5.setPrefWidth(100);
        lab6.setPrefWidth(100);
        lab7.setPrefWidth(100);
        lab8.setPrefWidth(100);
        lab9.setPrefWidth(100);
        lab10.setPrefWidth(100);
        listView0.setPrefWidth(100);
        
        
        root.getChildren().addAll(vsl,vusn,vname,vcat,vp1,vp2,vp3,vp4,vp5,vp6,vp7,vp8,vp9,vp10);


        
        
        Scene scene = new Scene(root, screenBounds.getWidth(), screenBounds.getHeight()-300);
        
        
        
        
        
        
        scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
        primaryStage.setMaximized(true);
        primaryStage.setScene(scene);
        primaryStage.show();
        primaryStage.close();
        
        
        
        // Bind the ListView scroll property
        Node n1 = listView1.lookup(".scroll-bar");
        
        if (n1 instanceof ScrollBar) {
            final ScrollBar bar1 = (ScrollBar) n1;
            Node n0 = listView0.lookup(".scroll-bar");
            Node n2 = listView2.lookup(".scroll-bar");
            Node n3 = listView3.lookup(".scroll-bar");
            
            Node n4 = lab1.lookup(".scroll-bar");
            Node n5 = lab2.lookup(".scroll-bar");
            Node n6 = lab3.lookup(".scroll-bar");
            Node n7 = lab4.lookup(".scroll-bar");
            Node n8 = lab5.lookup(".scroll-bar");
            Node n9 = lab6.lookup(".scroll-bar");
            Node n10 = lab7.lookup(".scroll-bar");
            Node n11 = lab8.lookup(".scroll-bar");
            Node n12 = lab9.lookup(".scroll-bar");
            Node n13 = lab10.lookup(".scroll-bar");
            
            System.out.println(n3);
            final ScrollBar bar3 = (ScrollBar) n3;
            final ScrollBar bar0 = (ScrollBar) n0;
            final ScrollBar bar4 = (ScrollBar) n4;
            final ScrollBar bar5 = (ScrollBar) n5;
            final ScrollBar bar6 = (ScrollBar) n6;
            final ScrollBar bar7 = (ScrollBar) n7;
            final ScrollBar bar8 = (ScrollBar) n8;
            final ScrollBar bar9 = (ScrollBar) n9;
            final ScrollBar bar10 = (ScrollBar) n10;
            final ScrollBar bar11 = (ScrollBar) n11;
            final ScrollBar bar12 = (ScrollBar) n12;
            final ScrollBar bar13 = (ScrollBar) n13;
            
            
            if (n2 instanceof ScrollBar) {
                final ScrollBar bar2 = (ScrollBar) n2;
                bar1.valueProperty().bindBidirectional(bar0.valueProperty());
                bar1.valueProperty().bindBidirectional(bar2.valueProperty());
                bar1.valueProperty().bindBidirectional(bar3.valueProperty());
                
                bar1.valueProperty().bindBidirectional(bar4.valueProperty());
                bar1.valueProperty().bindBidirectional(bar5.valueProperty());
                bar1.valueProperty().bindBidirectional(bar6.valueProperty());
                bar1.valueProperty().bindBidirectional(bar7.valueProperty());
                bar1.valueProperty().bindBidirectional(bar8.valueProperty());
                bar1.valueProperty().bindBidirectional(bar9.valueProperty());
                bar1.valueProperty().bindBidirectional(bar10.valueProperty());
                bar1.valueProperty().bindBidirectional(bar11.valueProperty());
                bar1.valueProperty().bindBidirectional(bar12.valueProperty());
                bar1.valueProperty().bindBidirectional(bar13.valueProperty());
            }
            
        }
        int width = (int) Math.round(screenBounds.getWidth());
        System.out.println(width);
        int size = (width/100)*60;
        
        System.out.println(size);
        root.setMaxWidth(size);
        root.setPrefHeight(screenBounds.getHeight());
        ap.getChildren().add(root);
        
        root.setMaxWidth(size);
		
		load(studdat.get(0), studdat.get(1));
		
	}
	
	
	public void load(String sem, String sec) throws IOException
	{
		InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+sem+sec+".xls");
		HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet sheet = workbook.getSheetAt(0);
        lap1= FXCollections.observableArrayList();
        lap2= FXCollections.observableArrayList();
        lap3= FXCollections.observableArrayList();
        lap4= FXCollections.observableArrayList();
        lap5= FXCollections.observableArrayList();
        lap6= FXCollections.observableArrayList();
        lap7= FXCollections.observableArrayList();
        lap8= FXCollections.observableArrayList();
        lap9= FXCollections.observableArrayList();
        lap10= FXCollections.observableArrayList();
        
        for(int i=0;i<usns.size()*4;i++)
        {
        	lap1.add(sheet.getRow(i).getCell(0).getStringCellValue());
        	lap2.add(sheet.getRow(i).getCell(1).getStringCellValue());
        	lap3.add(sheet.getRow(i).getCell(2).getStringCellValue());
        	lap4.add(sheet.getRow(i).getCell(3).getStringCellValue());
        	lap5.add(sheet.getRow(i).getCell(4).getStringCellValue());
        	lap6.add(sheet.getRow(i).getCell(5).getStringCellValue());
        	lap7.add(sheet.getRow(i).getCell(6).getStringCellValue());
        	lap8.add(sheet.getRow(i).getCell(7).getStringCellValue());
        	lap9.add(sheet.getRow(i).getCell(8).getStringCellValue());
        	lap10.add(sheet.getRow(i).getCell(9).getStringCellValue());
        }
        
        lab1.setItems(lap1);
        lab2.setItems(lap2);
        lab3.setItems(lap3);
        lab4.setItems(lap4);
        lab5.setItems(lap5);
        lab6.setItems(lap6);
        lab7.setItems(lap7);
        lab8.setItems(lap8);
        lab9.setItems(lap9);
        lab10.setItems(lap10);
        
        
        System.out.println(lap1);
        
	}
	
	
	public void importViva() throws IOException
	{
		FileChooser fileChooser = new FileChooser();
		File file = fileChooser.showOpenDialog(null);
		
		String str = file.getAbsolutePath().toString();
		InputStream ExcelFileToRead = new FileInputStream(str);
		
        Workbook w;
        try {
            w = Workbook.getWorkbook(ExcelFileToRead);
            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            // Loop over first 10 column and lines
            
             marks = new ArrayList<ArrayList<String>>();
             names = new ArrayList<String>();
            ArrayList<String> mark;
            int i = 6;
            int j = 1;
            
            
            
            System.out.println(sheet.getRows());
            while(j<sheet.getRows())
            {
            	i = 6;
            	names.add(sheet.getCell(0,j).getContents());
            	mark = new ArrayList<String>();
            	while(i<sheet.getColumns()-1)
            	{
            	
            	mark.add(sheet.getCell(i,j).getContents());
            	i++;
            	}
            	
            	marks.add(mark);
            	
            	j++;
            }
            System.out.println(names);
            System.out.println(marks);
        } 
        catch (BiffException em) {
            em.printStackTrace();
        }
        
       
        InputStream ExcelFileToRead1 = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+studdat.get(0)+studdat.get(1)+".xls");
		HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead1);
        HSSFSheet sheet = workbook.getSheetAt(0);
        
        System.out.println(sheet.getPhysicalNumberOfRows());
        
        
        System.out.println(marks.size() +"asda"+ marks.get(0).size()+ "adsdawf"+names.size());
        int k =1;
        for(int i=0;i<marks.size();i++)
        {
        	
        	for(int j=0;j<marks.get(0).size() && k<=sheet.getPhysicalNumberOfRows();j++)
        	{
        		sheet.getRow(k).createCell(j).setCellValue(marks.get(i).get(j));
        	}
        	System.out.println(k);
        	k=k+4;
        	
        }
        
        try {
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+studdat.get(0)+studdat.get(1)+".xls");
        workbook.write(fileOut);
        fileOut.close();
        }
        catch(Exception er)
        {
        	
        }
        
    	Alert alert = new Alert(AlertType.INFORMATION, "Saved as Document", ButtonType.OK);
        alert.getDialogPane().setMinHeight(Region.USE_PREF_SIZE);
        alert.show();
        
        load(studdat.get(0), studdat.get(1));
	}
	
	
	public void open_stud(ActionEvent e)
	{
		try {
			Stage primaryStage = new Stage();
			Parent root=FXMLLoader.load(getClass().getResource("Controller.fxml"));
			Screen screen = Screen.getPrimary();
			//BorderPane root = new BorderPane();
			
			
			
			Scene scene = new Scene(root,screenBounds.getWidth()/2, screenBounds.getHeight()/2);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.setTitle("Student Setup");
			primaryStage.show();
		} catch(Exception et) {
			et.printStackTrace();
		}
	}
	
	public void generateListEditable()
	{
		  
        lab1.setEditable(true);
        lab1.setCellFactory(TextFieldListCell.forListView());		
        lab1.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab1.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab1.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
        });
        
        
        
        lab2.setEditable(true);
        lab2.setCellFactory(TextFieldListCell.forListView());		
        lab2.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab2.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab2.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab3.setEditable(true);
        lab3.setCellFactory(TextFieldListCell.forListView());		
        lab3.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab3.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab3.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab4.setEditable(true);
        lab4.setCellFactory(TextFieldListCell.forListView());		
        lab4.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab4.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab4.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab5.setEditable(true);
        lab5.setCellFactory(TextFieldListCell.forListView());		
        lab5.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab5.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab5.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab6.setEditable(true);
        lab6.setCellFactory(TextFieldListCell.forListView());		
        lab6.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab6.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab6.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab7.setEditable(true);
        lab7.setCellFactory(TextFieldListCell.forListView());		
        lab7.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab7.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab7.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab8.setEditable(true);
        lab8.setCellFactory(TextFieldListCell.forListView());		
        lab8.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab8.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab8.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab9.setEditable(true);
        lab9.setCellFactory(TextFieldListCell.forListView());		
        lab9.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab9.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab9.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab10.setEditable(true);
        lab10.setCellFactory(TextFieldListCell.forListView());		
        lab10.setOnEditCommit(new EventHandler<ListView.EditEvent<String>>() {
			@Override
			public void handle(ListView.EditEvent<String> t) {
				lab10.getItems().set(t.getIndex(), t.getNewValue());
				System.out.println("setOnEditCommit");
			}
						
		});
        lab10.setOnEditCancel(new EventHandler<ListView.EditEvent<String>>() {
        	@Override
			public void handle(EditEvent<String> event) {
				// TODO Auto-generated method stub
				
			}
	
	});
        
        
        lab1.setFixedCellSize(24);
        lab2.setFixedCellSize(24);
        lab3.setFixedCellSize(24);
        lab4.setFixedCellSize(24);
        lab5.setFixedCellSize(24);
        lab6.setFixedCellSize(24);
        lab7.setFixedCellSize(24);
        lab8.setFixedCellSize(24);
        lab9.setFixedCellSize(24);
        lab10.setFixedCellSize(24);
       
	}
	
	
	public static class Person {
	   	 
        private final SimpleStringProperty usn;
    	private final SimpleStringProperty name;
        private final SimpleStringProperty classes;
        
        
        private Person(String string,String string1, String string2) {
        	this.usn = new SimpleStringProperty(string);
        	this.name =new SimpleStringProperty(string1);
            this.classes =new SimpleStringProperty(string2);
        }
        
        

		public String getUsn() {
            
            return usn.get();
        }
        
        public void setUsn(String u) {
           usn.set(u);
        }

        public String getName() {
            
            return name.get();
        }

        
        public void setName(String u) {
            name.set(u);
        }

        public String getClasses() {
        	
        	return classes.get();
        }

        public void setClasses(String u) {
            classes.set(u);
           
        }
       
	}
        
    	public void combine( ArrayList<ArrayList<String>> big) throws IOException
    	{
    		 int n;
    		 XWPFDocument docX2 = new XWPFDocument();
    	        
    	      XWPFParagraph paragraph = docX2.createParagraph();
    	      paragraph.setAlignment(ParagraphAlignment.CENTER);	      
    	      XWPFRun paragraphOneRunOne = paragraph.createRun();
    	      paragraphOneRunOne.setBold(true);
    	      paragraphOneRunOne.setText("DAYANANDA SAGAR COLLEGE OF ENGINEERING");
    	      paragraphOneRunOne.addBreak();
    	      
    	     
    	      XWPFRun paragraphOneRunTwo = paragraph.createRun();
    	      paragraphOneRunTwo.setBold(true);
    	      paragraphOneRunTwo.setText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING");
    	      paragraphOneRunTwo.addBreak();
    	      paragraphOneRunTwo.addBreak();
    	      
    	      
    	      XWPFRun paragraphOneRunThree = paragraph.createRun();
    	      paragraphOneRunThree.setBold(true);
    	      paragraphOneRunThree.setText("SECOND ATTENDANCE DISPLAY");
    	      paragraphOneRunThree.addBreak();
    	      
    	      XWPFRun paragraphOneRunFour = paragraph.createRun();
    	      paragraphOneRunFour.setBold(true);
    	      paragraphOneRunFour.setText("(Session: Jan 2019-May 2019)");
    	      paragraphOneRunFour.addBreak();
    	      paragraphOneRunFour.addBreak();
    	      paragraphOneRunFour.addBreak();
    	      
    	      XWPFParagraph paragraph1 = docX2.createParagraph();
    	      paragraph1.setAlignment(ParagraphAlignment.LEFT);
    	      
    	      XWPFRun paragraphTwoRunOne = paragraph1.createRun();
    	      paragraphTwoRunOne.setBold(true);
    	      paragraphTwoRunOne.setText("Class: 5th A                                                                                               Period: 16/8/2018 to 30/10/2018");
    	      
    	      
    	     
    	      
    	      
    	      //create table
    	      XWPFTable table = docX2.createTable();
    	      
    	      
    	      
    	      
    	      XWPFTableRow tableRowOne = table.createRow();
    	      table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1000));
    	      
    	      XWPFTableCell cell2=tableRowOne.getCell(0);
    	  	cell2.setText("Sl#");
    	  	CTTcPr tcpr = cell2.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge=tcpr.addNewVMerge();
    	  	vMerge.setVal(STMerge.RESTART); 
    	      
    	  	XWPFTableCell cell3=tableRowOne.createCell();
    	  	cell3.setText("USN");
    	  	CTTcPr tcpr1 = cell3.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge1=tcpr1.addNewVMerge();
    	  	vMerge1.setVal(STMerge.RESTART); 
    	      
    	  	
    	  	tableRowOne.createCell().setText("Subject ->");
    	  	
    	  	XWPFTableCell cell4=tableRowOne.createCell();
    	  	cell4.setText("ME");
    	  	CTTcPr tcpr2 = cell4.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge2=tcpr2.addNewHMerge();
    	  	vMerge2.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell5=tableRowOne.createCell();
    	  	CTTcPr tcpr3 = cell5.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge3=tcpr3.addNewHMerge();
    	  	vMerge3.setVal(STMerge.CONTINUE); 
    	  	
    	  	
    	  	XWPFTableCell cell6=tableRowOne.createCell();
    	  	cell6.setText("DBMS");
    	  	CTTcPr tcpr4 = cell6.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge4=tcpr4.addNewHMerge();
    	  	vMerge4.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell7=tableRowOne.createCell();
    	  	CTTcPr tcpr5 = cell7.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge5=tcpr5.addNewHMerge();
    	  	vMerge5.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell8=tableRowOne.createCell();
    	  	cell8.setText("SE");
    	  	CTTcPr tcpr6 = cell8.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge6=tcpr6.addNewHMerge();
    	  	vMerge6.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell9=tableRowOne.createCell();
    	  	CTTcPr tcpr7 = cell9.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge7=tcpr7.addNewHMerge();
    	  	vMerge7.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell10=tableRowOne.createCell();
    	  	cell10.setText("ATFL");
    	  	CTTcPr tcpr8 = cell10.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge8=tcpr8.addNewHMerge();
    	  	vMerge8.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell11=tableRowOne.createCell();
    	  	CTTcPr tcpr9 = cell11.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge9=tcpr9.addNewHMerge();
    	  	vMerge9.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell12=tableRowOne.createCell();
    	  	cell12.setText("AI");
    	  	CTTcPr tcpr10 = cell12.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge10=tcpr10.addNewHMerge();
    	  	vMerge10.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell13=tableRowOne.createCell();
    	  	CTTcPr tcpr11 = cell13.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge11=tcpr11.addNewHMerge();
    	  	vMerge11.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell14=tableRowOne.createCell();
    	  	cell14.setText("ADF");
    	  	CTTcPr tcpr12 = cell14.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge12=tcpr12.addNewHMerge();
    	  	vMerge12.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell15=tableRowOne.createCell();
    	  	CTTcPr tcpr13 = cell15.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge13=tcpr13.addNewHMerge();
    	  	vMerge13.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell16=tableRowOne.createCell();
    	  	cell16.setText("DBMS Lab");
    	  	CTTcPr tcpr14 = cell16.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge14=tcpr14.addNewHMerge();
    	  	vMerge14.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell17=tableRowOne.createCell();
    	  	CTTcPr tcpr15 = cell17.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge15=tcpr15.addNewHMerge();
    	  	vMerge15.setVal(STMerge.CONTINUE);
    	  	
    	  	XWPFTableCell cell18=tableRowOne.createCell();
    	  	cell18.setText("CN Lab");
    	  	CTTcPr tcpr16 = cell18.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge16=tcpr16.addNewHMerge();
    	  	vMerge16.setVal(STMerge.RESTART); 
    	  	
    	  	XWPFTableCell cell19=tableRowOne.createCell();
    	  	CTTcPr tcpr17 = cell19.getCTTc().addNewTcPr();
    	  	CTHMerge vMerge17=tcpr17.addNewHMerge();
    	  	vMerge17.setVal(STMerge.CONTINUE);
    	  	
    	  	
    	  	XWPFTableRow tableRowOne1 = table.createRow();
    	  	
    	  	XWPFTableCell cell20=tableRowOne1.getCell(0);
    	  	CTTcPr tcpr18 = cell20.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge18=tcpr18.addNewVMerge();
    	  	vMerge18.setVal(STMerge.CONTINUE); 
    	  	
    	  	XWPFTableCell cell21=tableRowOne1.createCell();
    	  	CTTcPr tcpr19 = cell21.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge19=tcpr19.addNewVMerge();
    	  	vMerge19.setVal(STMerge.CONTINUE); 
    	  	
    	  	
    	  	tableRowOne1.createCell().setText("Classes Conducted ->");
    	  	
    	  	for(int i = 0;i<8;i++)
    	  	{
    	  		XWPFTableCell cell22=tableRowOne1.createCell();
    		  	cell22.setText("30");
    		  	CTTcPr tcpr20 = cell22.getCTTc().addNewTcPr();
    		  	CTHMerge vMerge20=tcpr20.addNewHMerge();
    		  	vMerge20.setVal(STMerge.RESTART); 
    		  	
    		  	XWPFTableCell cell23=tableRowOne1.createCell();
    		  	CTTcPr tcpr21 = cell23.getCTTc().addNewTcPr();
    		  	CTHMerge vMerge21=tcpr21.addNewHMerge();
    		  	vMerge21.setVal(STMerge.CONTINUE);
    	  	}
    	  		
    	  	
    	  	XWPFTableRow tableRowOne2 = table.createRow();
    	  	
    	  	XWPFTableCell cell24=tableRowOne2.getCell(0);
    	  	CTTcPr tcpr21 = cell24.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge21=tcpr21.addNewVMerge();
    	  	vMerge21.setVal(STMerge.CONTINUE); 
    	  	
    	  	XWPFTableCell cell25=tableRowOne2.createCell();
    	  	CTTcPr tcpr22 = cell25.getCTTc().addNewTcPr();
    	  	CTVMerge vMerge22=tcpr22.addNewVMerge();
    	  	vMerge22.setVal(STMerge.CONTINUE); 
    	  	
    	  	tableRowOne2	.createCell().setText("Name");
    	  	
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      tableRowOne2.addNewTableCell().setText("A");
    	      tableRowOne2.addNewTableCell().setText("%");
    	      
    	     
    	      int[] cols = {20000,32000, 32000, 15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000,15000}; 
    		     
    	      for (int i = 0; i < table.getNumberOfRows(); i++) {
    	    	    XWPFTableRow row = table.getRow(i);
    	    	    int numCells = row.getTableCells().size();
    	    	    for (int j = 0; j < numCells; j++)
    	    	    {
    	    	        XWPFTableCell cell = row.getCell(j);
    	    	        CTTblWidth cellWidth = cell.getCTTc().addNewTcPr().addNewTcW();
    	    	        CTTcPr pr = cell.getCTTc().addNewTcPr();
    	    	        pr.addNewNoWrap();
    	    	        cellWidth.setW(BigInteger.valueOf(cols[j]));
    	    	        
    	    	        
    	    	    } 
    	    	}
    	       
    	     
    	     
    	      InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+".xls");
    	  	HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
    	  	HSSFSheet sheet = wb.getSheetAt(0);
    	  	HSSFRow row; 
    	  	
    	  	ArrayList<String> names = new ArrayList<String>();
    	  	ArrayList<String> usns = new ArrayList<String>();
    	  	
    	  	for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
    	  	{
    	  		
    	  	
    	  	String usn = sheet.getRow(i).getCell(0).toString();
    	  	String name = sheet.getRow(i).getCell(1).toString();
    	    
    	  	names.add(name);
    	  	usns.add(usn);
    	      
    	  	}

    	  	System.out.println(names);
    	  	
    	  	
    	  	
    	  	for(int i=0;i<names.size();i++)
    	  	{
    	  		table.createRow().getCell(0).setText(String.valueOf(i+1));
    	  		table.getRow(i+4).createCell().setText(usns.get(i));
    	  		table.getRow(i+4).createCell().setText(names.get(i));
    	  		
    	  		
    	  		for(int k=0;k<big.size();k++)
    	  		{
    	  			if(i<big.get(k).size())
        	  		{
    	  			table.getRow(i+4).createCell().setText(big.get(k).get(i).toString());
        	  		}
    	  		}
    	  	}
    	  	
    	  	
    	  	
    	      
    	  	FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+"consolidated.docx");
          docX2.write(fileOut);
          fileOut.close();
    		            
    	            System.out.println(".docx written successully");
    	              
    	}

 	
    	public void MarksCreate(ActionEvent e) throws IOException
    	{
    		BasicConfigurator.configure();
    	    	FileInputStream serviceAccount=
    		new FileInputStream("C:\\Users\\Mohan\\workspace\\fac\\serviceAccountKey.json");
    	FirebaseOptions options = new FirebaseOptions.Builder()
    			  .setCredentials(GoogleCredentials.fromStream(serviceAccount))
    			  .setDatabaseUrl("https://dsceapp-5ed7f.firebaseio.com")
    			  .build();

    FirebaseApp.initializeApp(options);
    ArrayList<DataSnapshot> Userlist = new ArrayList<DataSnapshot>();   
    try {
          final CountDownLatch latch1 = new CountDownLatch(1);
          DatabaseReference ref1= FirebaseDatabase.getInstance().getReference();
          DatabaseReference ref2;    
           ref2 = ref1.child("Marks/sem_2/");


           ref2.addListenerForSingleValueEvent(
        		   new ValueEventListener() {
            public void onDataChange(DataSnapshot dataSnapshot) {

                //ArrayList<Object> Userlist = new ArrayList<Object>();   
                ArrayList<ArrayList<String>> big_arr = new ArrayList<ArrayList<String>>();
               	                      for (DataSnapshot dsp : dataSnapshot.getChildren()) {
                      Userlist.add(dsp); 
                     
                    }
                //big_arr.add(Userlist);
                
     				 // System.out.println(Userlist.get(0)+"dsad"+Userlist.size());
     				     
                                            latch1.countDown();
    				     }

    		        	  public void onCancelled(DatabaseError error) {
    		        		  latch1.countDown();
    		        		  
    		        	  }
    		        	});
    		        	 latch1.await();
    			   } 
    			 catch (InterruptedException en) {
    			        en.printStackTrace();
    			    }
    			
    	ArrayList<ArrayList<String>> big = new ArrayList<ArrayList<String>>();
    	ArrayList<String> smol = new ArrayList<String>();
    	
    	for(DataSnapshot d: Userlist.get(0).getChildren())
    	{
    		smol = new ArrayList<String>();
    		FireData fir = d.getValue(FireData.class);
    		System.out.println(fir.getAtt());
    		smol.addAll(Arrays.asList(fir.getAtt().split(",")));
    		big.add(smol);
    		smol = new ArrayList<String>();
    		smol.addAll(Arrays.asList(fir.getPerc().split(",")));
    		big.add(smol);
    	}

    	System.out.println(big);
    		Marks(big);
    	}
    	
    	public void Marks(ArrayList<ArrayList<String>> big) throws IOException
    	{
    		       
    		XWPFDocument docX2 = new XWPFDocument();
    		 		 
    		      XWPFParagraph paragraph = docX2.createParagraph();
    		      paragraph.setAlignment(ParagraphAlignment.CENTER);	      
    		      XWPFRun paragraphOneRunOne = paragraph.createRun();
    		      paragraphOneRunOne.setBold(true);
    		      paragraphOneRunOne.setText("DAYANANDA SAGAR COLLEGE OF ENGINEERING");
    		      paragraphOneRunOne.addBreak();
    		      
    		     
    		      XWPFRun paragraphOneRunTwo = paragraph.createRun();
    	          paragraphOneRunTwo.setBold(true);
    		      paragraphOneRunTwo.setText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING");
    		      paragraphOneRunTwo.addBreak();
    		      paragraphOneRunTwo.addBreak();
    		      
    		      
    		      XWPFRun paragraphOneRunThree = paragraph.createRun();
    		      paragraphOneRunThree.setBold(true);
    		      paragraphOneRunThree.setText("SECOND TEST MARKS DISPLAY");
    		      paragraphOneRunThree.addBreak();
    		      
    		      XWPFRun paragraphOneRunFour = paragraph.createRun();
    		      paragraphOneRunFour.setBold(true);
    		      paragraphOneRunFour.setText("(Session: Jan 2019-May 2019)");
    		      paragraphOneRunFour.addBreak();
    		      paragraphOneRunFour.addBreak();
    		      paragraphOneRunFour.addBreak();
    		      
    		      XWPFParagraph paragraph1 = docX2.createParagraph();
    		      paragraph1.setAlignment(ParagraphAlignment.LEFT);
    		      
    		      XWPFRun paragraphTwoRunOne = paragraph1.createRun();
    		      paragraphTwoRunOne.setBold(true);
    		      paragraphTwoRunOne.setText("Class: 5th A                                                                                                                                    Max. Marks:10");
    		      
    		     
    		      //create table
    		      XWPFTable table = docX2.createTable();
//    		      table.setWidth(3*1440);

    		      //create first row
    		      XWPFTableRow tableRowOne = table.getRow(0);
    		      table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1000));
    		      tableRowOne.getCell(0).setText("SL.#");
    		      tableRowOne.addNewTableCell().setText("USN");
    		      tableRowOne.addNewTableCell().setText("Subject");
    		      tableRowOne.addNewTableCell().setText("ME");
    		      tableRowOne.addNewTableCell().setText("CN");
    		      tableRowOne.addNewTableCell().setText("DBMS");
    		      tableRowOne.addNewTableCell().setText("SE");
    		      tableRowOne.addNewTableCell().setText("ATFL");
    		      tableRowOne.addNewTableCell().setText("AIAT");
    		      tableRowOne.addNewTableCell().setText("ADF");
//    		      tableRowOne.addNewTableCell().setText("P7");
//    		      tableRowOne.addNewTableCell().setText("P8");
//    		      tableRowOne.addNewTableCell().setText("P9");
//    		      tableRowOne.addNewTableCell().setText("P10");
//    		      tableRowOne.addNewTableCell().setText("P11");
//    		      tableRowOne.addNewTableCell().setText("P12");
    		    
    		     
    		   int[] cols = {8000,20000, 20000, 10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000}; 
    		     
    		      for (int i = 0; i < table.getNumberOfRows(); i++) {
    		    	    XWPFTableRow row = table.getRow(i);
    		    	    int numCells = row.getTableCells().size();
    		    	    for (int j = 0; j < numCells; j++)
    		    	    {
    		    	        XWPFTableCell cell = row.getCell(j);
    		    	        CTTblWidth cellWidth = cell.getCTTc().addNewTcPr().addNewTcW();
    		    	        CTTcPr pr = cell.getCTTc().addNewTcPr();
    		    	        pr.addNewNoWrap();
    		    	        cellWidth.setW(BigInteger.valueOf(cols[j]));
    		    	        
    		    	        
    		    	    } 
    		    	}
    		      
    		      InputStream ExcelFileToRead = new FileInputStream("D:\\Book1.xls");
    			  	HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
    			  	HSSFSheet sheet = wb.getSheetAt(0);
    			  	HSSFRow row; 
    			  	
    			  	ArrayList<String> names = new ArrayList<String>();
    			  	ArrayList<String> usns = new ArrayList<String>();
    			  	
    			  	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++)
    			  	{
    			  		
    			  	
    			  	String usn = sheet.getRow(i).getCell(0).toString();
    			  	String name = sheet.getRow(i).getCell(1).toString();
    			    
    			  	names.add(name);
    			  	usns.add(usn);
    			      
    			  	}
    			  	
    			  	for(int i=0;i<names.size();i++)
    			  	{
    			  		table.createRow().getCell(0).setText(String.valueOf(i+1));
    			  		table.getRow(i+1).getCell(1).setText(usns.get(i));
    			  		table.getRow(i+1).getCell(2).setText(names.get(i));
    			  		System.out.println(i);
    			  		if(i<big.get(0).size())
    			  		{
    			  		for(int k=0;k<big.size();k++)
    			  		{
    			  			table.getRow(i+1).getCell(k+3).setText(big.get(k).get(i).toString());
    			  		}
    			  		}
    			  	}
    			  	
    		/*System.out.println("Enter no. of students");
    		int n= Integer.parseInt(in.readLine());
    		//ChangeOrientation obj= new ChangeOrientation();

    		for( int i=0;i<n;i++)
    		      {XWPFTableRow tableRowNext = table.createRow();
    		      
    		      tableRowNext.getCell(0).setText(Integer.toString(i));
    		      tableRowNext.getCell(1).setText("idk");
    		      tableRowNext.getCell(2).setText("idk");
    		      tableRowNext.getCell(3).setText("idk");
    		      tableRowNext.getCell(4).setText("idk");
    		      tableRowNext.getCell(5).setText("idk");
    		      tableRowNext.getCell(6).setText("idk");
    		      tableRowNext.getCell(7).setText("idk");
    		      tableRowNext.getCell(8).setText("idk");
    		      tableRowNext.getCell(9).setText("idk");
    		      }
    		      
    		      

    */
    		FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\mark.docx");
    	      docX2.write(fileOut);
    	      fileOut.close();
    			            
    		            System.out.println(".docx written successully");
    		              
    	}
        
        
        
    	public static class Person_Marks {
    	  	 
    	    private final SimpleStringProperty usn;
    		private final SimpleStringProperty name;
    	    private final SimpleStringProperty marks;
    	    
    	    
    	    private Person_Marks(String string,String string1, String string2) {
    	    	this.usn = new SimpleStringProperty(string);
    	    	this.name =new SimpleStringProperty(string1);
    	        this.marks =new SimpleStringProperty(string2);
    	    }
    	    
    	    

    		public String getUsn() {
    	        
    	        return usn.get();
    	    }
    	    
    	    public void setUsn(String u) {
    	       usn.set(u);
    	    }

    	    public String getName() {
    	        
    	        return name.get();
    	    }

    	    
    	    public void setName(String u) {
    	        name.set(u);
    	    }

    	    public String getMarks() {
    	    	
    	    	return marks.get();
    	    }

    	    public void setMarks(String u) {
    	        marks.set(u);
    	       
    	    }
    	}

    	

    
    
	
	
}
