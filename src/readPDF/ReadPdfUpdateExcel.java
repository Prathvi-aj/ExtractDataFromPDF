package readPDF;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.*;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;


public class ReadPdfUpdateExcel {
	static Connection con=null;
	
	
	public static void main(String args[]) throws IOException {
		
		
		Fillo fillo = new Fillo();	
		Recordset recordset=null;
		List<String> invoiceLink=new ArrayList<>();
		
		try {
			
			con= fillo.getConnection(System.getProperty("user.dir")+"/Abhishek_assigment.xlsx");
			recordset=con.executeQuery("select * from Sheet1");
			
			try {
			con.createTable("Invoice_Detail",new String[] {"SN","Order_Number","Invoice_Number","Buyer_Name_Address","Order_Date","Invoice_Date"});
			}catch(Exception e) {
				System.out.println("Sheet already present.");
			}
			
			while(recordset.next()) {
				invoiceLink.add(recordset.getField("Invoice Download Link"));
			}
			
			
		} catch (FilloException e) {
			e.printStackTrace();
		}
		
		
		for(int i=1;i<invoiceLink.size();i++) {
		getPDFFromURL(invoiceLink.get(i),i);
		ReadPDFandUpdateexcel("invoice"+i+".pdf",i);
		}
		
		
		
	}
	
	public static void getPDFFromURL(String url1,int i){
		URL url=null;
		byte[] ba1 = new byte[1024];
	    int baLength=0;;
	    FileOutputStream fos1=null;
		//System.out.print(url1);
		try {
		
			url = new URL(url1);
		fos1 = new FileOutputStream("invoice"+i+".pdf");
		  
		//System.out.print("Connecting to " + url.toString());
	      URLConnection urlConn = url.openConnection();

		        try {

	          InputStream is1 = url.openStream();
	          while ((baLength = is1.read(ba1)) != -1) {
	              fos1.write(ba1, 0, baLength);
	          }
	          fos1.flush();
	          fos1.close();
	          is1.close();
		       
		        }catch(Exception e) {
		        	
		        e.printStackTrace();	
		        }
	}catch(Exception e) {
		e.printStackTrace();
	}
	}
	
	
	public static void ReadPDFandUpdateexcel(String pdfName,int i) {
		
		
		try {
			
		PDDocument document =PDDocument.load(new File(System.getProperty("user.dir")+"/"+pdfName));
		PDFTextStripper stripper = new PDFTextStripper();
		String text=stripper.getText(document);
		//System.out.println(text+"  ++");
		document.close();
		
		String[][] replacements = {{"Order Number", "ON"}, 
                {"Invoice Number", "IN"},{"Order Date", "OD"}, 
                {"Invoice Date", "ID"},{"SHIP TO:", "ST"}};
		
		String text1 = text;
		for(String[] replacement: replacements) {
			text1 = text1.replace(replacement[0], replacement[1]);
		}

		
		String order_no=text1.substring(text1.indexOf("ON")+3,text1.indexOf("IN"));
		String invoice_no=text1.substring(text1.indexOf("IN")+3,text1.indexOf("ST"));
		String user_name_address=text1.substring(text1.indexOf("ST")+3,text1.indexOf("ID"));
		String order_date=text1.substring(text1.indexOf("ID")+3,text1.indexOf("OD"));
		String invoice_date=text1.substring(text1.indexOf("OD")+3,text1.indexOf("SN"));
		String insert_query="insert into invoice_detail(SN,Order_Number,Invoice_Number,Buyer_Name_Address,Order_Date,Invoice_Date) values('"+i+"','"+order_no+"','"+invoice_no+"','"+user_name_address+"','"+order_date+"','"+invoice_date+"')";
		//System.out.println(insert_query);
		con.executeUpdate(insert_query);
		
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		
		}	
	
	
}