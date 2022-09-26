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
import com.spire.pdf.PdfDocument;
import com.spire.pdf.utilities.PdfTable;
import com.spire.pdf.utilities.PdfTableExtractor;


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
			con.createTable("Invoice_Detail",new String[] {"SN","Order_Number","Invoice_Number","Buyer_Name_Address","Order_Date","Invoice_Date","PRODUCT_TITLE","HSN","TAXABLE_VALUE","DISCOUNT","TAX_RATE_AND_CATEGORY","TOTAL"});
			}catch(Exception e) {
				System.out.println("Sheet already present.");
			}
			
			while(recordset.next()) {
				invoiceLink.add(recordset.getField("Invoice Download Link"));
			}
			
			
		} catch (FilloException e) {
			e.printStackTrace();
		}
		
		int i=1;
		for(i=1;i<invoiceLink.size();i++) {
		getPDFFromURL(invoiceLink.get(i),i);
		ReadPDFandUpdateexcel("invoice"+i+".pdf",i);
		
		}
		
		ReadPDFandUpdateexcel("jorwp231840.pdf",++i);
		ReadPDFandUpdateexcel("jorwp231864.pdf",++i);
		ReadPDFandUpdateexcel("jorwp231884.pdf",++i);
		
	
		
		
		
		
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
	
	
	
	public static void ReadPDFandUpdateexcel(String pdfName,int n) {
		
		
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
		
		
		PdfDocument pdf = new PdfDocument(System.getProperty("user.dir")+"/"+pdfName);
        PdfTableExtractor extractor = new PdfTableExtractor(pdf);
        
        ArrayList<String> tableDataField=new ArrayList<>();
        HashMap<Integer,HashMap<String,String>> tableData=new HashMap<>();
        
        for (int pageIndex = 0; pageIndex < pdf.getPages().getCount(); pageIndex++) {
            PdfTable[] tableLists = extractor.extractTable(pageIndex);
            if (tableLists != null && tableLists.length > 0) {
                for (PdfTable table : tableLists) {
                    for (int i = 0; i < 3; i++) {
                    	if(i==0) {
                        for (int j = 0; j < table.getColumnCount(); j++) {
                        	
                        	tableDataField.add(table.getText(i, j));
                            
                        }
                    	}
                    	
                        if(i!=0) {
                        	HashMap<String,String> tablecolumne=new HashMap<>();
                            
                        for (int j = 0; j < table.getColumnCount(); j++) {
                        	
                        	tablecolumne.put(tableDataField.get(j),table.getText(i, j));
                        	
                        	
                        }
                        tableData.put(i,tablecolumne);
                        }
                    
                    }
                   
                    
                }
            }
        }
        
       // System.out.println(tableData);
 
		
		
		
        HashMap<String , String> tableData1=new HashMap<>();
        HashMap<String , String> tableData2=new HashMap<>();
		
        String insert_query="";
        double tax=0;
		for(int k=1;k<=tableData.size();k++) {
			
			
			
			
			if(k==1) {
				tableData1=tableData.get(k);
				insert_query="insert into invoice_detail(SN,Order_Number,Invoice_Number,Buyer_Name_Address,Order_Date,Invoice_Date,PRODUCT_TITLE,HSN,TAXABLE_VALUE,DISCOUNT,TAX_RATE_AND_CATEGORY,TOTAL) values('"+n+"','"+order_no+"','"+invoice_no+"','"+user_name_address+"','"+order_date+"','"+invoice_date+"','"+tableData1.get("Description")+"','"+tableData1.get("HSN")+"','"+(tableData1.get("Taxes")).substring(16)+"','"+tableData1.get("Discount")+"','"+(tableData1.get("Taxes")).substring(0,5)+"','"+tableData1.get("Total")+"')";
				
			}
			else if (k==2 && !(tableData.get(k).get("HSN").equals(""))){
				tableData2=tableData.get(k);
				
				for(Map.Entry<String , String> e: tableData2.entrySet()) {
					if(e.getKey().startsWith("Taxable"))
					tax+=Double.parseDouble(e.getValue().substring(3));
				}
				for(Map.Entry<String , String> e: tableData1.entrySet()) {
					if(e.getKey().startsWith("Taxable"))
					tax+=Double.parseDouble(e.getValue().substring(3));
				}
				double total=Double.parseDouble((tableData1.get("Total")).substring(3))+Double.parseDouble((tableData2.get("Total")).substring(3));
				//System.out.println(tax);
				insert_query="insert into invoice_detail(SN,Order_Number,Invoice_Number,Buyer_Name_Address,Order_Date,Invoice_Date,PRODUCT_TITLE,HSN,TAXABLE_VALUE,DISCOUNT,TAX_RATE_AND_CATEGORY,TOTAL) values('"+n+"','"+order_no+"','"+invoice_no+"','"+user_name_address+"','"+order_date+"','"+invoice_date+"','"+tableData1.get("Description")+"','"+tableData1.get("HSN")+"','"+tax+"','"+tableData1.get("Discount")+"','"+(tableData1.get("Taxes")).substring(0,5)+"','"+total+"')";
				
			}
				
		}
		
		//System.out.println(insert_query);
		con.executeUpdate(insert_query);
		
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		
		}	
	

	
	
	
	
	
	
}