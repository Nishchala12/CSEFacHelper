package application;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;






public class ChangeOrientation{
 public void edit_t(int n,String usn, String name, XWPFTable table2, ArrayList C, ArrayList V, ArrayList R, ArrayList T) throws IOException
{
//	XWPFDocument docX2=new XWPFDocument();
	
	
	XWPFTableRow row2=table2.createRow();
	
	
	XWPFTableCell cell2=row2.getCell(0);
	cell2.setText(Integer.toString(n));
	CTTcPr tcpr = cell2.getCTTc().addNewTcPr();
	CTVMerge vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.RESTART); 
	
	
	cell2=row2.getCell(1);
	cell2.setText(usn);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.RESTART);
	 
	cell2=row2.getCell(2);
	cell2.setText(name);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.RESTART);
	
	cell2=row2.getCell(3);
	cell2.setText("C");
	
	System.out.println(C.size()+"sdaa");
	for(int i=0;i<C.size();i++)
	{
	cell2=row2.getCell(i+4);
	cell2.setText(C.get(i).toString());
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.RESTART);
	}
	
	/*cell2=row2.getCell(4);
	cell2.setText("04");
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.RESTART);
	*/
	row2 = table2.createRow();
	
	
	cell2=row2.getCell(0);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(1);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(2);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(3);
	cell2.setText("R");
	
	for(int i=0;i<C.size();i++)
	{
	cell2=row2.getCell(i+4);
	cell2.setText(V.get(i).toString());
	}
	/*cell2=row2.getCell(4);
	cell2.setText("44");*/
	
    row2 = table2.createRow();
	
	
	cell2=row2.getCell(0);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(1);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(2);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	
	cell2=row2.getCell(3);
	cell2.setText("V");
	
	for(int i=0;i<C.size();i++)
	{
	cell2=row2.getCell(i+4);
	cell2.setText(R.get(i).toString());
	}
	/*cell2=row2.getCell(4);
	cell2.setText("66");
	*/
row2 = table2.createRow();
	
	
	cell2=row2.getCell(0);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(1);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(2);
	tcpr = cell2.getCTTc().addNewTcPr();
	vMerge=tcpr.addNewVMerge();
	vMerge.setVal(STMerge.CONTINUE);
	
	cell2=row2.getCell(3);
	cell2.setText("T");
	
	for(int i=0;i<C.size();i++)
	{
	cell2=row2.getCell(i+4);
	cell2.setText(T.get(i).toString());
	}
	/*cell2=row2.getCell(4);
	cell2.setText("88");
	*/
	
	
	int[] cols = {8000,15000, 15000, 10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000,10000}; 
    
    for (int i = 0; i < table2.getNumberOfRows(); i++) {
  	    XWPFTableRow row1 = table2.getRow(i);
  	    int numCells = row1.getTableCells().size();
  	    for (int j = 0; j < numCells; j++)
  	    {
  	        XWPFTableCell cell = row1.getCell(j);
  	        CTTblWidth cellWidth = cell.getCTTc().addNewTcPr().addNewTcW();
  	        CTTcPr pr = cell.getCTTc().addNewTcPr();
  	        pr.addNewNoWrap();
  	        cellWidth.setW(BigInteger.valueOf(cols[j]));
  	        
//  	        table.getRow(i).setHeight(1000);
  	        
  	    } 
  	}
	
	
//	FileOutputStream fileOut = new FileOutputStream("C:\\Users\\oindr\\SDM\\new.docx");
//	docX2.write(fileOut);
}
 
 public void populateDoc(XWPFDocument docX2, List usn, List names, List p1, List p2,  List p3, List p4, List p5, List p6, List p7, List p8, List p9, List p10) throws IOException
 {
	 XWPFTable table2 = docX2.createTable();
		
		XWPFTableRow row = table2.getRow(0); // First row  
	    // Columns  
	    row.getCell(0).setText("Sl. No.");  
	    row.addNewTableCell().setText("USN");  
	    row.addNewTableCell().setText("NAME");
	    row.addNewTableCell().setText("  ");
	    row.addNewTableCell().setText("P1");
	    row.addNewTableCell().setText("P2");
	    row.addNewTableCell().setText("P3");
	    row.addNewTableCell().setText("P4");
	    row.addNewTableCell().setText("P5");
	    row.addNewTableCell().setText("P6");
	    row.addNewTableCell().setText("P7");
	    row.addNewTableCell().setText("P8");
	    row.addNewTableCell().setText("P9");
	    row.addNewTableCell().setText("P10");
	    row.addNewTableCell().setText("P11");
	    row.addNewTableCell().setText("P12");
		
	    int j=0;
	    for(int i=0;i<usn.size();i++)
	    {
	    	ArrayList C = new ArrayList<>();
	    	ArrayList V = new ArrayList<>();
	    	ArrayList R = new ArrayList<>();
	    	ArrayList T = new ArrayList<>();
	    	C.addAll(Arrays.asList(p1.get(j),p2.get(j), p3.get(j), p4.get(j), p5.get(j),p6.get(j),p7.get(j),p8.get(j),p9.get(j),p10.get(j)));
	    	j++;
	    	V.addAll(Arrays.asList(p1.get(j),p2.get(j), p3.get(j), p4.get(j), p5.get(j),p6.get(j),p7.get(j),p8.get(j),p9.get(j),p10.get(j)));
	    	j++;
	    	R.addAll(Arrays.asList(p1.get(j),p2.get(j), p3.get(j), p4.get(j), p5.get(j),p6.get(j),p7.get(j),p8.get(j),p9.get(j),p10.get(j)));
	    	j++;
	    	T.addAll(Arrays.asList(p1.get(j),p2.get(j), p3.get(j), p4.get(j), p5.get(j),p6.get(j),p7.get(j),p8.get(j),p9.get(j),p10.get(j)));
	    	j++;
	    	
	    	
	    	edit_t(i+1, usn.get(i).toString(), names.get(i).toString(), table2, C, V, R, T);
	    
	    }
	    
	    
 }


 
 
}