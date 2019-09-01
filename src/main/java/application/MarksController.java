package application;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Optional;
import java.util.concurrent.CountDownLatch;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import com.google.auth.oauth2.GoogleCredentials;
import com.google.firebase.FirebaseApp;
import com.google.firebase.FirebaseOptions;
import com.google.firebase.database.DataSnapshot;
import com.google.firebase.database.DatabaseError;
import com.google.firebase.database.DatabaseReference;
import com.google.firebase.database.FirebaseDatabase;
import com.google.firebase.database.ValueEventListener;

import application.Attendance.Person;
import application.DashBoardController.Person_Marks;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.TableColumn.CellEditEvent;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.stage.Screen;
import javafx.util.Pair;

public class MarksController 
{
	public static String in="INSTRUCTIONS  FOR USAGE:\r\n" + 
			"\r\n" + 
			"  #TO LOAD A SPREADSHEET: \r\n" + 
			"  1) CLICK THE LOAD SPREADSHEET\r\n" + 
			"     BUTTON\r\n" + 
			"  2) LOAD THE REQUIRED FILE FROM \r\n" + 
			"     THE SYSTEM EXPLORER\r\n" + 
			"\r\n" + 
			"\r\n" + 
			"  #TO SAVE SPREADSHEET:\r\n" + 
			"  1) CLICK THE SAVE SPREADSHEET\r\n" + 
			"     BUTTON AFTER ENTERING THE\r\n" + 
			"     REQUIRED DETAILS \r\n"+	
			"\r\n" + 
			"\r\n" + 
			"  #TO SAVE AND SYNC\r\n" + 
			"     ONLINE:\r\n" + 
			"  1) CLICK THE SAVE BUTTON. THE  \r\n" + 
			"     DATA GETS SYNCED TO FIREBASE\r\n" + 
			"  2) CLICK THE  SYNC\r\n" + 
			"     BUTTON TO TRANSFER THE DATA \r\n" + 
			"     INTO THE OFFICIAL WORD \r\n" + 
			"     DOCUMENT\r\n" + 
			"\r\n" + 
			"";
	
	String tfsem = "";
	String tfsec = "";

	@FXML
	TextField semester=new TextField();
	@FXML
    TextField section=new TextField();
	@FXML
	TitledPane tpma;
	@FXML 
	AnchorPane ap_marks, ap;
	@FXML
	Label in2;
	@FXML
	String rootpath = "C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM";
	@FXML
	VBox utility_marks;
	String coursecode ="";
	Boolean exp=false;
	
	@FXML
	private TableView<Person_Marks> table_marks;
	
	TableColumn usnCol2,nameCol2,cie1Col,cie2Col,cie3Col,aatCol,asCol,newCol,avgCol;
	
	 private ObservableList<Person_Marks> data1 =
		        FXCollections.observableArrayList();
	@FXML
	Button savefiremarksbtn, loadfiremarksbtn, loadmarksbtn, savemarksbtn;
	
	 
    ArrayList<ArrayList<String>> marks;
	ArrayList<String> names;
	
	 Rectangle2D screenBounds = Screen.getPrimary().getBounds();
	    ArrayList<String> usns = new ArrayList<>();
	    ArrayList<String> name = new ArrayList<>();
	    static ArrayList<String> studdat = new ArrayList<String>();
	    
	    
	    public void initialize() throws IOException
		{		
			
			
			 Rectangle2D screenBounds = Screen.getPrimary().getBounds();
		       //header.setPrefWidth(screenBounds.getWidth());
		       utility_marks.setPrefHeight(screenBounds.getHeight());
		      // vbox_nav.setPrefHeight(screenBounds.getHeight());
		        
		       	in2.setText(in);
		       	in2.setWrapText(true);
		        in2.setStyle("-fx-font-family: \"Times New Roman \"; -fx-font-size: 20; -fx-text-fill: white; -fx-background-color:#6464a5;");
		        int width = (int) screenBounds.getWidth();
		        int utilsize = (width/100)*20;
		        int lab_size = (width/100)*60;
		        utility_marks.setPrefWidth(utilsize);
		        
		        table_marks.setPrefWidth(lab_size);
		        table_marks.setPrefHeight(screenBounds.getHeight());
		        utility_marks.setPrefHeight(screenBounds.getHeight());
		        ap_marks.setPrefWidth(lab_size);
		        ap_marks.setPrefHeight(screenBounds.getHeight());

		        
		        loadfiremarksbtn.setPrefWidth(utilsize);
		        savefiremarksbtn.setPrefWidth(utilsize);
		        savemarksbtn.setPrefWidth(utilsize);
		        loadmarksbtn.setPrefWidth(utilsize);
		        
		        table_marks.setEditable(true);
		        usnCol2 = new TableColumn("USN");
		        usnCol2.setMinWidth(100);
		        usnCol2.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("usn"));
		       
		        nameCol2 = new TableColumn("NAME");
		        nameCol2.setMinWidth(100);
		        nameCol2.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("name"));

		        cie1Col = new TableColumn("CIE-1");
		        cie1Col.setMinWidth(100);
		        cie1Col.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("cie1"));
		        cie1Col.setCellFactory(TextFieldTableCell.forTableColumn());
		        cie1Col.setOnEditCommit(
		        new EventHandler<CellEditEvent<Person_Marks, String>>() {
		        @Override
		        public void handle(CellEditEvent<Person_Marks, String> t) {
		        ((Person_Marks) t.getTableView().getItems().get(
		        t.getTablePosition().getRow())
		        ).setCie1(t.getNewValue());
		        }
		        }
		        );

		        cie2Col = new TableColumn("CIE-2");
		        cie2Col.setMinWidth(100);
		        cie2Col.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("cie2"));
		        cie2Col.setCellFactory(TextFieldTableCell.forTableColumn());
		        cie2Col.setOnEditCommit(
		        new EventHandler<CellEditEvent<Person_Marks, String>>() {
		        @Override
		        public void handle(CellEditEvent<Person_Marks, String> t) {
		        ((Person_Marks) t.getTableView().getItems().get(
		        t.getTablePosition().getRow())
		        ).setCie2(t.getNewValue());
		        }
		        }
		        );

		        cie3Col = new TableColumn("CIE-3");
		        cie3Col.setMinWidth(100);
		        cie3Col.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("cie3"));
		        cie3Col.setCellFactory(TextFieldTableCell.forTableColumn());
		        cie3Col.setOnEditCommit(
		        new EventHandler<CellEditEvent<Person_Marks, String>>() {
		        @Override
		        public void handle(CellEditEvent<Person_Marks, String> t) {
		        ((Person_Marks) t.getTableView().getItems().get(
		        t.getTablePosition().getRow())
		        ).setCie3(t.getNewValue());
		        }
		        }
		        );

		        aatCol = new TableColumn("AAT");
		        aatCol.setMinWidth(100);
		        aatCol.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("aat"));
		        aatCol.setCellFactory(TextFieldTableCell.forTableColumn());
		        aatCol.setOnEditCommit(
		        new EventHandler<CellEditEvent<Person_Marks, String>>() {
		        @Override
		        public void handle(CellEditEvent<Person_Marks, String> t) {
		        ((Person_Marks) t.getTableView().getItems().get(
		        t.getTablePosition().getRow())
		        ).setAat(t.getNewValue());
		        }
		        }
		        );

		        asCol = new TableColumn("ASSGMT");
		        asCol.setMinWidth(100);
		        asCol.setCellValueFactory(
		                new PropertyValueFactory<Person_Marks, String>("as"));
		        asCol.setCellFactory(TextFieldTableCell.forTableColumn());
		        asCol.setOnEditCommit(
		        new EventHandler<CellEditEvent<Person_Marks, String>>() {
		        @Override
		        public void handle(CellEditEvent<Person_Marks, String> t) {
		        ((Person_Marks) t.getTableView().getItems().get(
		        t.getTablePosition().getRow())
		        ).setAs(t.getNewValue());
		        }
		        }
		        );
		        avgCol = new TableColumn("TOTAL");
		        avgCol.setMinWidth(100);
		        avgCol.setCellValueFactory(
		                new PropertyValueFactory<Person, String>("avg"));
		        
		        tpma.setExpanded(false);
		        usnCol2.setVisible(false);
		        nameCol2.setVisible(false);
		        cie1Col.setVisible(false);
		        cie2Col.setVisible(false);
		        cie3Col.setVisible(false);
		        aatCol.setVisible(false);
		        asCol.setVisible(false);
		       // newCol.setVisible(false);
		        avgCol.setVisible(false);
		        
		        table_marks.setItems(data1);
		        table_marks.getColumns().addAll(usnCol2, nameCol2, cie1Col, cie2Col, cie3Col, aatCol, asCol, avgCol);
		}
	    
	    public void openMarks(ActionEvent e)
	    {
	    	if(!exp==true)
	    	{
	    	exp =true;
	    	tpma.setExpanded(true);
	    	}
	    	else
	    	{
	    		exp=false;
	    		tpma.setExpanded(false);
	    	}
	    }
	    public  void LoadMarks(ActionEvent e)throws IOException
		{
	    	 studdat.clear();
			    tfsem = semester.getText().toString();
			    tfsem = tfsem.toUpperCase();
			    tfsec = section.getText().toString();
			    tfsec = tfsec.toUpperCase();
			    studdat.add(tfsem);
			    studdat.add(tfsec);

			    usnCol2.setVisible(true);
		        nameCol2.setVisible(true);
		        cie1Col.setVisible(true);
		        cie2Col.setVisible(true);
		        cie3Col.setVisible(true);
		        aatCol.setVisible(true);
		        asCol.setVisible(true);
		        avgCol.setVisible(true);
		    table_marks.getColumns().clear();
		    table_marks.getColumns().addAll(usnCol2, nameCol2, cie1Col, cie2Col, cie3Col, aatCol, asCol, avgCol);
		    table_marks.setItems(data1);
			
			InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+studdat.get(0)+studdat.get(1)+".xls");
			HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
			HSSFSheet sheet = wb.getSheetAt(0);
			data1.clear();
			table_marks.setItems(data1);
			HSSFRow row; 
			HSSFCell cell;
			String [] sheetrows;
			Iterator rows = sheet.rowIterator();
			Iterator<Row> x=sheet.iterator();
			Row next=x.next();
			
			sheetrows = new String[8];
			
			int k =1;
			
			System.out.println(sheet.getRow(k).getCell(0));
			
			while(k<sheet.getPhysicalNumberOfRows())
			{
				data1.add(new Person_Marks(sheet.getRow(k).getCell(0).getStringCellValue(),
						sheet.getRow(k).getCell(1).getStringCellValue(),
						sheet.getRow(k).getCell(2).getStringCellValue(),
						sheet.getRow(k).getCell(3).getStringCellValue(),
						sheet.getRow(k).getCell(4).getStringCellValue(),
						sheet.getRow(k).getCell(5).getStringCellValue(),
						sheet.getRow(k).getCell(6).getStringCellValue(),
						sheet.getRow(k).getCell(7).getStringCellValue()));
				System.out.println(sheet.getRow(k).getCell(2).getStringCellValue());
				k++;
			}
			table_marks.setItems(data1);
			
			/*for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
			{
			
			HSSFRow row1 = sheet.getRow(i);
			Iterator cells = row1.cellIterator();
			while (cells.hasNext())
			{
			
			cells.next();
			}
			String usn = sheet.getRow(i).getCell(0).toString();
			String name = sheet.getRow(i).getCell(1).toString();
			String c1 = sheet.getRow(i).getCell(2).toString();
			String c2 = sheet.getRow(i).getCell(3).toString();
			String c3 = sheet.getRow(i).getCell(4).toString();
			String aat = sheet.getRow(i).getCell(5).toString();
			String as = sheet.getRow(i).getCell(6).toString();
			String avg = sheet.getRow(i).getCell(7).toString();
			
			
			
			data1.add(new Person_Marks(usn,name,c1, c2, c3, aat, as, avg));*/
			
				
		
			
			/*table_marks.getColumns().clear();
			table_marks.getColumns().addAll(usnCol2, nameCol2, cie1Col, cie2Col, cie3Col, aatCol, asCol, avgCol);
			table_marks.setItems(data1);*/
			
		}
		
		public  void SaveMarks(ActionEvent e)throws IOException
		{
			InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+studdat.get(0)+studdat.get(1)+".xls");
			HSSFWorkbook  workbook = new HSSFWorkbook(ExcelFileToRead);
	        HSSFSheet spreadsheet = workbook.getSheetAt(0);

	        
		    
	        int i=1;
	        for(Person_Marks dsce: data1)
			{
	        	float sum=0;
				if(!dsce.getCie1().equals(""))
				{
					spreadsheet.getRow(i).createCell(2).setCellValue(dsce.getCie1());
				System.out.println(dsce.getCie1());
				sum+= (Integer.parseInt((dsce.getCie1().toString()))*0.2);
				spreadsheet.getRow(i).createCell(7).setCellValue(sum+"");
				}
				if(!dsce.getCie2().equals(""))
				{
					spreadsheet.getRow(i).createCell(3).setCellValue(dsce.getCie2());
				System.out.println(dsce.getCie2());
				sum+= (Integer.parseInt((dsce.getCie2().toString()))*0.2);
				spreadsheet.getRow(i).createCell(7).setCellValue(sum+"");
				}
				if(!dsce.getCie3().equals(""))
				{
					spreadsheet.getRow(i).createCell(4).setCellValue(dsce.getCie3());
				System.out.println(dsce.getCie3());
				sum+= (Integer.parseInt((dsce.getCie3().toString()))*0.2);
				spreadsheet.getRow(i).createCell(7).setCellValue(sum+"");
				}
				if(!dsce.getAat().equals(""))
				{
					spreadsheet.getRow(i).createCell(5).setCellValue(dsce.getAat());
				System.out.println(dsce.getAat());
				sum+= Integer.parseInt((dsce.getAat().toString()));
				spreadsheet.getRow(i).createCell(7).setCellValue(sum+"");
				}
				if(!dsce.getAs().equals(""))
				{
					spreadsheet.getRow(i).createCell(6).setCellValue(dsce.getAs());
				System.out.println(dsce.getAs());
				sum+= Integer.parseInt((dsce.getAs().toString()));
				spreadsheet.getRow(i).createCell(7).setCellValue(sum+"");
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

		public void MarksCreate(ActionEvent e) throws IOException
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
    	    private final SimpleStringProperty cie1;
    	    private final SimpleStringProperty cie2;
    	    private final SimpleStringProperty cie3;
    	    private final SimpleStringProperty aat;
    	    private final SimpleStringProperty as;
    	    private final SimpleStringProperty avg;
    	    
    	    
    	    private Person_Marks(String string,String string1, String string2, String string3, String string4, String string5, String string6, String string7) {
    	    	this.usn = new SimpleStringProperty(string);
    	    	this.name =new SimpleStringProperty(string1);
    	        this.cie1 =new SimpleStringProperty(string2);
    	        this.cie2 =new SimpleStringProperty(string3);
    	        this.cie3 =new SimpleStringProperty(string4);
    	        this.aat =new SimpleStringProperty(string5);
    	        this.as =new SimpleStringProperty(string6);
    	        this.avg =new SimpleStringProperty(string7);
    	        
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

    	    public String getCie1() {
    	    	
    	    	return cie1.get();
    	    }

    	    public void setCie1(String u) {
    	        cie1.set(u);
    	       
    	    }
    	    public String getCie2() {
    	    	
    	    	return cie2.get();
    	    }

    	    public void setCie2(String u) {
    	        cie2.set(u);
    	       
    	    }
    	    public String getCie3() {
    	    	
    	    	return cie3.get();
    	    }

    	    public void setCie3(String u) {
    	        cie3.set(u);
    	       
    	    }
    	    public String getAat() {
    	    	
    	    	return aat.get();
    	    }

    	    public void setAat(String u) {
    	        aat.set(u);
    	       
    	    }
    	    public String getAs() {
    	    	
    	    	return as.get();
    	    }

    	    public void setAs(String u) {
    	        as.set(u);
    	       
    	    }
    	    public String getAvg() {
    	    	
    	    	return avg.get();
    	    }

    	    public void setAvg(String u) {
    	        avg.set(u);
    	       
    	    }
    	}


	    
}
