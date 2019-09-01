package application;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.ProtocolException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.CountDownLatch;

import org.apache.http.HttpEntity;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.TextUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONException;
import org.json.JSONObject;
import org.omg.CORBA.portable.InputStream;

import com.google.api.client.http.HttpResponse;
import com.google.firebase.auth.FirebaseAuth;
import com.google.firebase.auth.FirebaseAuthException;
import com.google.firebase.auth.UserRecord;
import com.google.firebase.auth.UserRecord.CreateRequest;
import com.google.firebase.database.DatabaseReference;
import com.google.firebase.database.FirebaseDatabase;
import com.google.firebase.internal.NonNull;
import com.sun.javafx.tk.Toolkit.Task;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.Event;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TextField;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.image.ImageView;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

public class LoginController {
	 
    @FXML
    private TextField u_name;
    
    @FXML
	ImageView exit, exit1,exit2;
    @FXML
    private PasswordField password;

    @FXML
    private Button login;
    
    @FXML
    private Button SignUp;
    
    @FXML
    private Button SignIn;
    
    @FXML
	ComboBox c=new ComboBox();
    
    @FXML
    private TextField username;

    @FXML
    private PasswordField password2;

    @FXML
    private TextField section;
    
    @FXML
    private TextField name;
    String e;
    String p;
    String uname;
    String pass;
    String result;
    
    
    public void screenDestroy(Event e)
	{
		Stage primstage = (Stage) exit.getScene().getWindow();
		primstage.close();
	}
    
    public void screenDestroy1(Event e)
   	{
   		Stage primstage = (Stage) exit1.getScene().getWindow();
   		primstage.close();
   	}
    public void screenDestroy2(Event e)
   	{
   		Stage primstage = (Stage) exit2.getScene().getWindow();
   		primstage.close();
   	}
    
    @FXML
    private void initialize()
    {
    	
    
    	 ObservableList<Integer> options = 
    	       	    FXCollections.observableArrayList(1,2,3,4,5,6,7);
    	       
    		        c.getItems().addAll(options);
    	
    	
    }
    
    @FXML
    void SignUp(ActionEvent event) throws IOException {
    	
  
    	Stage primaryStage=new Stage();
    	primaryStage.initStyle(StageStyle.UNDECORATED);
		Parent root=FXMLLoader.load(getClass().getResource("Login.fxml"));
		Scene scene = new Scene(root,800,600);
		scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
		primaryStage.setScene(scene);
		primaryStage.show();
		Stage primstage = (Stage) login.getScene().getWindow();
		primstage.close();
		
    }
    
    @FXML
    void Login(ActionEvent event) throws IOException, FirebaseAuthException, InterruptedException {
    	String u=username.getText();
    	String p=password2.getText();
    	
    	
    	 if(username.getText().isEmpty()) {
 	    	Alert alert = new Alert(AlertType.ERROR);
 	    	alert.setTitle("Error Dialog");
 	   // 	alert.setHeaderText("Look, an Error Dialog");
 	    	alert.setContentText("Please enter your username ");
 	    	alert.showAndWait();
 	    	return;
 	    	 }
    	 else if(password2.getText().isEmpty()) {
	 	    	Alert alert = new Alert(AlertType.ERROR);
		    	alert.setTitle("Error Dialog");
		  //  	alert.setHeaderText("Look, an Error Dialog");
		    	alert.setContentText("Please enter your Password");
		    	alert.showAndWait();
		    	return;
	    	 }
    	 else
    	 {
    		
    		 loginUser(u,p);
    		 
    		 
    	 }
    	 
    	 
    }
    
    
    
    void loginUser(String email, String password) throws FirebaseAuthException, InterruptedException, IOException
    {
    	
    	 uname= username.getText();
    	 pass=password2.getText();
    	 
    	  	 System.out.println("name"+ uname);
    	//  	System.out.println(e);
    	  	validate();
    	  
    	  	
    	  	 if(result=="true")
    	  	 {
    	 UserRecord userRecord = FirebaseAuth.getInstance().getUserByEmail(email);
    	   
    	   
    	 
    	
    	 System.out.println("Successfully fetched user data: " + userRecord.getEmail());
    	// UserRecord userRecord = FirebaseAuth.getInstance().getUserByPassword(password);
    	// System.out.println("Successfully fetched user data: " + userRecord.getPassword());
    	 System.out.println("name"+ uname);
    	String us=userRecord.getEmail();
    	 
    	if(uname.equals(us)==true ) 
    	{
    		Stage primaryStage=new Stage();
    		primaryStage.initStyle(StageStyle.UNDECORATED);
    		Parent root=FXMLLoader.load(getClass().getResource("FrontEnd.fxml"));
    		Scene scene = new Scene(root,800,700);
    		scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
    		primaryStage.setScene(scene);
    		
    		primaryStage.show();
    	}
    	
    	
    	
    	  	 }
    	
 
}; 
    		 	
    
    
    
    
    void validate() throws IOException 
    {

	URL url = new URL ("https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=AIzaSyBypTF-ja1z39DB3S6wGmjgyoC0h7AWv_U");

HttpURLConnection con = (HttpURLConnection)url.openConnection();
try {
	con.setRequestMethod("POST");
} catch (ProtocolException e1) {
	// TODO Auto-generated catch block
	e1.printStackTrace();
}
con.setRequestProperty("Content-Type", "application/json; utf-8");
con.setRequestProperty("Accept", "application/json");
con.setDoOutput(true);
String jsonInputString = "{\"email\":\""+uname+"\",\"password\":\""+pass+"\"}";
//JSONObject json=new JSONObject();
try(OutputStream os = con.getOutputStream()) {
    byte[] input = jsonInputString.getBytes("utf-8");
    os.write(input, 0, input.length);           
} catch (IOException e) {
	// TODO Auto-generated catch block
	System.out.println("12");
	Alert alert = new Alert(AlertType.ERROR);
	alert.setTitle("Error Dialog");
//  	alert.setHeaderText("Look, an Error Dialog");
	alert.setContentText("The entered Email ID / Password is incorrect");
	alert.showAndWait(); 
	return;
}

try(BufferedReader br = new BufferedReader(
		  new InputStreamReader(con.getInputStream(), "utf-8"))) {
		    StringBuilder response = new StringBuilder();
		    String responseLine = null;
		    while ((responseLine = br.readLine()) != null) {
		        response.append(responseLine.trim());
		    }
		    try
	          { 
		    		System.out.println(response.getClass());
	        	  JSONObject jsonObject=new JSONObject(response.toString());
	        	  System.out.println(jsonObject.get("registered"));
	        	 result=jsonObject.get("registered").toString(); 
	          }
		    catch(JSONException e)
		    {
		    	System.out.println("error");
		    	
		    	
		    }
		    System.out.println(response.toString());
			//result=json.getString("registered");
    	  	
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			System.out.println("123");
			Alert alert = new Alert(AlertType.ERROR);
	    	alert.setTitle("Error Dialog");
	  //  	alert.setHeaderText("Look, an Error Dialog");
	    	alert.setContentText("The entered email ID / Password is incorrect");
	    	alert.showAndWait();
	    	return;
		} 
          


catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("1234");
			Alert alert = new Alert(AlertType.ERROR);
	    	alert.setTitle("Error Dialog");
	  //  	alert.setHeaderText("Look, an Error Dialog");
	    	alert.setContentText("The entered email ID / Password is incorrect");
	    	alert.showAndWait(); 
	    	return;
		}

    
	Stage primstage = (Stage) SignIn.getScene().getWindow();
	primstage.close();

    
 }
    
    
    
    
    
    @SuppressWarnings("unused")
	void registerUser(String email, String password) throws FirebaseAuthException
    {
    	CreateRequest request = new CreateRequest()
			    .setEmail(email)
			    .setEmailVerified(true)
			    .setPassword(password)
			    .setDisabled(false);
    	UserRecord userRecord = FirebaseAuth.getInstance().createUser(request);
    	System.out.println("Successfully created new user: " + userRecord.getUid());
    	 
    	final CountDownLatch latch1 = new CountDownLatch(1);
    	 String tchr_name=name.getText();
    	 String sec= section.getText().toUpperCase();
    	 String sem=c.getSelectionModel().getSelectedItem().toString();

    	DatabaseReference ref1 = FirebaseDatabase.getInstance().getReference("user/"+userRecord.getUid());
    	 DatabaseReference child_name = FirebaseDatabase.getInstance().getReference();
		 	 child_name=ref1.child("email");
		 	child_name.setValueAsync(userRecord.getEmail());
		 	 DatabaseReference child_descp = FirebaseDatabase.getInstance().getReference();
	    	 child_descp=ref1.child("section");
	    	 child_descp.setValueAsync(sec);
	    	 DatabaseReference child_answer = FirebaseDatabase.getInstance().getReference();
	    	 child_answer=ref1.child("teacher name");
	    	 child_answer.setValueAsync(tchr_name);
	    	 DatabaseReference child_sem = FirebaseDatabase.getInstance().getReference();
	    	 child_sem=ref1.child("Semester");
	    	 child_sem.setValueAsync(sem);


			/*try {
				UserRecord userRecord1 = FirebaseAuth.getInstance().createUser(request);
				System.out.println("Created....");
			} catch (FirebaseAuthException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
    }
    
    @FXML
    void SignUp2(ActionEvent event) throws IOException, FirebaseAuthException {
    	
    	
    	  e=u_name.getText();
    	 String tchr_name=name.getText();
    //	 String e=email.getText();
    	  p=password.getText();
    	 
    	 if(p.length()<6)
    	 {
    		 Alert alert = new Alert(AlertType.ERROR);
		    	alert.setTitle("Error Dialog");
		  //  	alert.setHeaderText("Look, an Error Dialog");
		    	alert.setContentText("Password should be 6 characters long");

		    	alert.showAndWait(); 
	    	 }
    	 
  
    		Workbook workbook = new HSSFWorkbook();
	        Sheet spreadsheet = workbook.createSheet("sample2");

	        Row row = spreadsheet.createRow(0);

	        for (int j = 0; j<1; j++) {
	         
	            	row.createCell(j).setCellValue("Nmae");
	           row.createCell(j+1).setCellValue("EMAIL");
	            row.createCell(j+2).setCellValue("Password");
	      
	            
	        }

	        for (int i = 0; i <1; i++) {
	            row = spreadsheet.createRow(i+1);
	            for (int j = 0; j < 1; j++) {
	            	 row.createCell(j).setCellValue(tchr_name); 
	           	 row.createCell(j+1).setCellValue(e); 
	            	 row.createCell(j+2).setCellValue(p); 
	          
	                }
	         
	            }
	   FileOutputStream fileOut = new FileOutputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\FacInfo\\Workbook2.xls");
	        workbook.write(fileOut);
	        fileOut.close();
    	
    	 if(u_name.getText().isEmpty()) {
    	Alert alert = new Alert(AlertType.ERROR);
    	alert.setTitle("Error Dialog");
    	alert.setHeaderText("Look, an Error Dialog");
    	alert.setContentText("Please enter your username ");

    	alert.showAndWait();
    	 }
  
    	
    	 else if( password.getText().isEmpty()) {
	 	    	Alert alert = new Alert(AlertType.ERROR);
		    	alert.setTitle("Error Dialog");
		    	alert.setHeaderText("Look, an Error Dialog");
		    	alert.setContentText("Please confirm the password");

		    	alert.showAndWait(); 

    	 
    	 }
    	 
    	 
    	 
    	 else
    	 {
    		 	registerUser(e, p);
    		 	Stage primaryStage=new Stage();
    		 	primaryStage.initStyle(StageStyle.UNDECORATED);
				Parent root=FXMLLoader.load(getClass().getResource("SignUp.fxml.fxml"));
				Scene scene = new Scene(root,800,600);
				scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
				primaryStage.setScene(scene);
				primaryStage.show();
    	 }
    	 
    	 
    	 
    	 Stage primstage = (Stage) SignUp.getScene().getWindow();
    		primstage.close();
    	    
    	 
    	 

    }
    
    
    
    

}
