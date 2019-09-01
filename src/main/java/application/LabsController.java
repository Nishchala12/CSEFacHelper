package application;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ListView.EditEvent;
import javafx.scene.control.cell.TextFieldListCell;
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

public class LabsController
{
	public static String in= "\r\n" + 
			"\r\n" + 
			" INSTRUCTIONS  FOR USAGE:\r\n" + 
			"\r\n" + 
			" #TO LOAD A SPREADSHEET: \r\n" + 
			" 1) CLICK THE LOAD SPREADSHEET\r\n" + 
			"     BUTTON\r\n" + 
			" 2) ENTER THE SEMESTER, SECTION,\r\n" + 
			"     AND BATCH DATA\r\n" + 
			" 3) LOAD THE REQUIRED FILE FROM \r\n" + 
			"    THE SYSTEM EXPLORER\r\n" + 
			"\r\n" + 
			"\r\n" + 
			" #TO IMPORT VIVA FROM MOODLE:\r\n" + 
			" 1) CLICK THE SAVE SPREADSHEET\r\n" + 
			"     BUTTON\r\n" + 
			" 2) LOAD THE REQUIRED FILE FROM \r\n" + 
			"     THE SYSTEM EXPLORER\r\n" + 
			"\r\n" + 
			"\r\n" + 
			" #TO SAVE AND CONSOLIDATE \r\n" + 
			"    ONLINE:\r\n" + 
			" 1) CLICK THE SAVE BUTTON. THE  \r\n" + 
			"     DATA GETS SYNCED TO FIREBASE\r\n" + 
			" 2) CLICK THE  CONSOLIDATE \r\n" + 
			"      BUTTON TO TRANSFER THE DATA \r\n" + 
			"      INTO THE OFFICIAL WORD \r\n" + 
			"      DOCUMENT\r\n" ;
	@FXML
	Label in3 ;
	@FXML 
	AnchorPane ap, ap_attendance, header, ap_lab,ap_marks;
	@FXML
	VBox utility_lab, utility_attend, utility_marks;
	
	String rootpath = "C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM";
	
	@FXML
	Button savespbtn,loadspbtn, savefir, syncsave, loadlab, savelab, importMoodle,saveFirLab, loadFirLab, savefiremarksbtn, loadfiremarksbtn, loadmarksbtn, savemarksbtn;
	
	ArrayList<ArrayList<String>> marks;
	ArrayList<String> names;
	Stage primaryStage = new Stage();
	HBox root = new HBox();
	static ArrayList<String> studdat = new ArrayList<String>();
	
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
    String tfsem = "";
	String tfsec = "";

	@FXML
	TextField semester=new TextField();
	@FXML
    TextField section=new TextField();
	@FXML
    TitledPane tpla;
	Boolean exp=false;
    @FXML
    public void initialize()
    {
    	in3.setText(in);
    	in3.setWrapText(true);
        in3.setStyle("-fx-font-family: \"Times New Roman \"; -fx-font-size: 20; -fx-text-fill: white; -fx-background-color:#6464a5;");
    	 Rectangle2D screenBounds = Screen.getPrimary().getBounds();
	       //header.setPrefWidth(screenBounds.getWidth());
    	 
    	// ap_lab.setPrefWidth(screenBounds.getWidth());
	     //   ap_lab.setPrefHeight(screenBounds.getHeight());
	       utility_lab.setPrefHeight(screenBounds.getHeight());
	      // vbox_nav.setPrefHeight(screenBounds.getHeight());
	        
	        
	        int width = (int) screenBounds.getWidth();
	        int utilsize = (width/100)*20;
	        int lab_size = (width/100)*60;
	        utility_lab.setPrefWidth(utilsize);
	        ap.setPrefWidth(lab_size);
	        ap.setPrefHeight(screenBounds.getHeight());
	        
	        tpla.setExpanded(false);
	        
	        loadlab.setPrefWidth(utilsize);
	        savelab.setPrefWidth(utilsize);
	        importMoodle.setPrefWidth(utilsize);
	        loadFirLab.setPrefWidth(utilsize);
	        saveFirLab.setPrefWidth(utilsize);
	        
	       
    }
    
    
    public void openLabs(ActionEvent e)
    {
    	if(!exp==true)
    	{
    	exp =true;
    	tpla.setExpanded(true);
    	}
    	else
    	{
    		exp=false;
    		tpla.setExpanded(false);
    	}
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
		
		
		studdat.clear();
	    tfsem = semester.getText().toString();
	    tfsem = tfsem.toUpperCase();
	    tfsec = section.getText().toString();
	    tfsec = tfsec.toUpperCase();
	    studdat.add(tfsem);
	    studdat.add(tfsec);

   
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
			dataarr1.add("R");
			dataarr1.add("V");
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
        int k =2;
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
    
	
}
