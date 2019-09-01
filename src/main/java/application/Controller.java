package application;


import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.CountDownLatch;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.google.firebase.database.DataSnapshot;
import com.google.firebase.database.DatabaseError;
import com.google.firebase.database.DatabaseReference;
import com.google.firebase.database.FirebaseDatabase;
import com.google.firebase.database.ValueEventListener;

import application.Attendance.Person;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.TableColumn.CellEditEvent;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.stage.Stage;

public class Controller {
	
	@FXML
	 private TableView<Person> table = new TableView<Person>();
	@FXML
	ComboBox c=new ComboBox();
	@FXML
	TextField text=new TextField();
	@FXML
	TextField cname;
	@FXML
	Button load,Save;
	
	
	 private ObservableList<Person> data;
	 TableColumn USNCol, NameCol;
	 ArrayList<String> name = new ArrayList<String>();
		ArrayList<String> usns = new ArrayList<String>();
		List<String> name1 = new ArrayList<>();
		List<String> usn = new ArrayList<>();
		PrintWriter fileloc = null;
		String output,tt,file, cnamestr;
		
	 
	 
	@FXML
	public void initialize() throws IOException
	{
		table.setEditable(true);
		
		  data=FXCollections.observableArrayList(
         new Person("",""));


        USNCol = new TableColumn("USN");
       USNCol.setMinWidth(100);
       USNCol.setCellValueFactory(
           new PropertyValueFactory<Person, String>("USN"));
       USNCol.setCellFactory(TextFieldTableCell.forTableColumn());
       USNCol.setOnEditCommit(
           new EventHandler<CellEditEvent<Person, String>>() {
               @Override
               public void handle(CellEditEvent<Person, String> t) {
                   ((Person) t.getTableView().getItems().get(
                           t.getTablePosition().getRow())
                           ).setUSN(t.getNewValue());
               }
           }
       );


        NameCol = new TableColumn("Name");
       NameCol.setMinWidth(100);
       NameCol.setCellValueFactory(
           new PropertyValueFactory<Person, String>("name"));
       NameCol.setCellFactory(TextFieldTableCell.forTableColumn());
       NameCol.setOnEditCommit(
           new EventHandler<CellEditEvent<Person, String>>() {
               @Override
               public void handle(CellEditEvent<Person, String> t) {
                   ((Person) t.getTableView().getItems().get(
                       t.getTablePosition().getRow())
                       ).setName(t.getNewValue());
               }
           }
       );
       
       
       table.setItems(data);
       table.getColumns().addAll(USNCol, NameCol);
       
       ObservableList<Integer> options = 
       	    FXCollections.observableArrayList(1,2,3,4,5,6,7);
       
	        c.getItems().addAll(options);
	}
	
	public void Load(ActionEvent e)throws IOException
	{
		table.getColumns().clear();
		table.getColumns().addAll(USNCol, NameCol);
		table.setItems(data);
		
		output = c.getSelectionModel().getSelectedItem().toString();
		tt=text.getText();
		file=output.toUpperCase()+tt.toUpperCase();
		
		cnamestr = cname.getText();

		
		FileChooser fileChooser = new FileChooser();
		File file = fileChooser.showOpenDialog(null);
		
		String str = file.getAbsolutePath().toString();
		
	InputStream ExcelFileToRead = new FileInputStream(str);
	HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
	HSSFSheet sheet = wb.getSheetAt(0);
	HSSFRow row; 
	data.clear();
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
	System.out.println(usn + name);
	data.add(new Person(usn,name));
		
	}
	}
	
	public void showSubs(ActionEvent e)throws IOException
	{
		output = c.getSelectionModel().getSelectedItem().toString();
        
        try {
            final CountDownLatch latch1 = new CountDownLatch(1);
            DatabaseReference ref= FirebaseDatabase.getInstance().getReference().child("Subjects/");


             ref.addListenerForSingleValueEvent(
          		new ValueEventListener() {
	              public void onDataChange(DataSnapshot d) {
	            	  if(d.hasChild(output))
	            		  cname.setText(d.child(output).getValue().toString());
	            	  else
	            		  cname.setText("");
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
	}
	
	public void Save(ActionEvent e)throws IOException
	{
		output = c.getSelectionModel().getSelectedItem().toString();
		tt=text.getText();
		file=output.toUpperCase()+tt.toUpperCase();
		
		cnamestr = cname.getText().toString();
		cnamestr = cnamestr.toUpperCase();
		
	    try {
	            final CountDownLatch latch1 = new CountDownLatch(1);
	            DatabaseReference ref= FirebaseDatabase.getInstance().getReference().child("Subjects");
	        	 DatabaseReference child_name = FirebaseDatabase.getInstance().getReference();
	        	 	
	        	 child_name=ref.child(output);
	        	 child_name.setValueAsync(cnamestr);
	        	 latch1.countDown();
	        	 
	        	System.out.println("Succesfull");
	        	 
	        	latch1.await();
	    	} 
	    catch (InterruptedException ef) {
	    			        ef.printStackTrace();
	    }
		
		
		TableColumn<Person, String> column1 = NameCol; // column you want

		
		for (Person item1 : table.getItems()) {
		    name1.add((String) column1.getCellObservableValue(item1).getValue());
		}
		
		TableColumn<Person, String> column2 = USNCol; // column you want

		
		for (Person item2 : table.getItems()) {
		    usn.add((String)column2.getCellObservableValue(item2).getValue());
		}
	
	        fileloc = new PrintWriter("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\StudentData\\"+file+".txt");

		try(BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\StudentData\\"+file+".txt",true))) 
		{  
			
			for(int i=0;i<usn.size();i++)
			{
		    bufferedWriter.write(usn.get(i)+"\t"+name1.get(i)+"\t");
		    bufferedWriter.newLine();
			}
		    bufferedWriter.close();
		    
		    HSSFWorkbook workbook = new HSSFWorkbook();
	        HSSFSheet sheet = workbook.createSheet();
	        
	        int p1 = 0;
	        for(int i=0;i<usn.size();i++)
	        {
	        	p1=p1+4;
	        }
	        
	        
		    for(int i=0;i<p1;i++)
	        {
	        	sheet.createRow(i).createCell(0).setCellValue("");
	        	sheet.getRow(i).createCell(1).setCellValue("");
	        	sheet.getRow(i).createCell(2).setCellValue("");
	        	sheet.getRow(i).createCell(3).setCellValue("");
	        	sheet.getRow(i).createCell(4).setCellValue("");
	        	sheet.getRow(i).createCell(5).setCellValue("");
	        	sheet.getRow(i).createCell(6).setCellValue("");
	        	sheet.getRow(i).createCell(7).setCellValue("");
	        	sheet.getRow(i).createCell(8).setCellValue("");
	        	sheet.getRow(i).createCell(9).setCellValue("");
	        	
	        }
		    
		    try {
		        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Lab\\"+file+".xls");
		        workbook.write(fileOut);
		        fileOut.close();
		        }
		        catch(Exception er)
		        {
		        	
		        }
		    
		    HSSFWorkbook workbook_at = new HSSFWorkbook();
	        HSSFSheet sheet_at = workbook_at.createSheet();
	        
	        sheet_at.createRow(0).createCell(0).setCellValue("		DAYANANDA SAGAR COLLEGE OF ENGINEERING		");
	        sheet_at.createRow(1).createCell(0).setCellValue("	DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING	");
	        sheet_at.createRow(2).createCell(0).setCellValue("		DATE: ");
	        sheet_at.createRow(3).createCell(0).setCellValue("		SUBJECT: ");
	        sheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 0, 3));
	        sheet.addMergedRegionUnsafe(new CellRangeAddress(3, 3, 0, 3));
	        sheet_at.createRow(4).createCell(0).setCellValue("USN");
	        sheet_at.getRow(4).createCell(1).setCellValue("Name");
	        sheet_at.getRow(4).createCell(2).setCellValue("Classes");
	        sheet_at.getRow(4).createCell(3).setCellValue("Percentage");
	        
	        System.out.println(usn.size()+"dsadad"+name1.size());
	        
		    for(int i=5;i<usn.size()+5;i++)
		    {
		    	sheet_at.createRow(i).createCell(0).setCellValue(usn.get(i-5));
		    	sheet_at.getRow(i).createCell(1).setCellValue(name1.get(i-5));
		    	sheet_at.getRow(i).createCell(2).setCellValue("");
		    	sheet_at.getRow(i).createCell(3).setCellValue("");
		    	System.out.println(i);
		    }
		    System.out.println(sheet_at.getPhysicalNumberOfRows());
		    try {
		        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+file+".xls");
		        workbook_at.write(fileOut);
		        fileOut.close();
		        }
		        catch(Exception er)
		        {
		        	
		        }
		    
		    
		    HSSFWorkbook workbook_mk = new HSSFWorkbook();
	        HSSFSheet sheet_mk = workbook_mk.createSheet();
	        
	        
	        sheet_mk.createRow(0).createCell(0).setCellValue("USN");
	        sheet_mk.getRow(0).createCell(1).setCellValue("Name");
	        sheet_mk.getRow(0).createCell(2).setCellValue("CIE-1");
	        sheet_mk.getRow(0).createCell(3).setCellValue("CIE-2");
	        sheet_mk.getRow(0).createCell(4).setCellValue("CIE-3");
	        sheet_mk.getRow(0).createCell(5).setCellValue("AAT");
	        sheet_mk.getRow(0).createCell(6).setCellValue("ASSGMT");
	        sheet_mk.getRow(0).createCell(7).setCellValue("TOTAL");
	        
	        System.out.println(usn.size()+"dsadad"+name1.size());
	        
		    for(int i=1;i<usn.size();i++)
		    {
		    	sheet_mk.createRow(i).createCell(0).setCellValue(usn.get(i-1));
		    	sheet_mk.getRow(i).createCell(1).setCellValue(name1.get(i-1));
		    	sheet_mk.getRow(i).createCell(2).setCellValue("");
		    	sheet_mk.getRow(i).createCell(3).setCellValue("");
		    	sheet_mk.getRow(i).createCell(4).setCellValue("");
		    	sheet_mk.getRow(i).createCell(5).setCellValue("");
		    	sheet_mk.getRow(i).createCell(6).setCellValue("");
		    	sheet_mk.getRow(i).createCell(7).setCellValue("");
		    	System.out.println(i);
		    }
		    System.out.println(sheet_mk.getPhysicalNumberOfRows());
		    try {
		        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Marks\\"+file+".xls");
		        workbook_mk.write(fileOut);
		        fileOut.close();
		        }
		        catch(Exception er)
		        {
		        	
		        }
	        
		    
		   
		    try {
		        BufferedWriter writer = new BufferedWriter(new FileWriter("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\FacInfo\\facinfo.txt"));
		        writer.flush();
		        writer.write(cnamestr);
		      //  writer.newLine();
				//writer.write(ccodestr);
				writer.close();
		    } catch (IOException ex) {
		        System.out.println("File could not be created");
		    }
	        
		}
		
		name1.clear();
		usn.clear();
		
		
		 Alert alert=new Alert(AlertType.INFORMATION);
	        alert.setTitle("Information Dialog");
	        alert.setHeaderText(null);
	        alert.setContentText("Data Saved!");
	        alert.showAndWait();
			return;
	}
	
	public static class Person {
		 
        private final SimpleStringProperty USN;
        private final SimpleStringProperty Name;
 
        Person(String usn, String name)
        {
            this.USN = new SimpleStringProperty(usn);
            this.Name = new SimpleStringProperty(name);
           
        }
 
        public String getUSN() {
            return USN.get();
        }
 
        public void setUSN(String usn) {
            USN.set(usn);
        }
 
        public String getName() {
            return Name.get();
        }
 
        public void setName(String usn) {
            Name.set(usn);
        }
 
       
    }
	
}






