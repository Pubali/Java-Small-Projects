package com.cassandra.practice.excelops;

import com.datastax.driver.core.Cluster;
import com.datastax.driver.core.Session;
import com.datastax.driver.core.ResultSet;
import java.io.FileInputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InsertExcelFileData {
    public static void main( String [] args ) throws Exception
    {
        String fileName="G:\\book.xlsx";
        Vector dataHolder=read(fileName);
        saveToDatabase(dataHolder);
    }
    public static Vector read(String fileName)    {
        Vector cellVectorHolder = new Vector();
        try{
            FileInputStream myInput = new FileInputStream(fileName);
            //POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
            XSSFWorkbook myWorkBook = new XSSFWorkbook(myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator rowIter = mySheet.rowIterator();
            while(rowIter.hasNext()){
                XSSFRow myRow = (XSSFRow) rowIter.next();
                Iterator cellIter = myRow.cellIterator();
                //Vector cellStoreVector=new Vector();
                List list = new ArrayList();
                while(cellIter.hasNext()){
                    XSSFCell myCell = (XSSFCell) cellIter.next();
                    list.add(myCell);
                }
                cellVectorHolder.addElement(list);
            }
        }catch (Exception e){e.printStackTrace(); }
        return cellVectorHolder;
    }
    private static void saveToDatabase(Vector dataHolder) {
        String ClientAdd="";
        String Page="";
        String AccessDate="";
        String   ProcessTime="";
        String Bytes="";
        System.out.println(dataHolder);

        for(Iterator iterator = dataHolder.iterator();iterator.hasNext();) {
            List list = (List) iterator.next();
            ClientAdd = list.get(0).toString();
            Page = list.get(1).toString();
            AccessDate = list.get(2).toString();
            ProcessTime = list.get(3).toString();

            try {
               /* Class.forName("com.mysql.jdbc.Driver").newInstance();
                Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/test", "root", "welcome");
                System.out.println("connection made...");
                PreparedStatement stmt=con.prepareStatement("INSERT INTO ClickStream(ClientAdd,Page,AccessDate,ProcessTime) VALUES(?,?,?,?)");
                stmt.setString(1, ClientAdd);
                stmt.setString(2, Page);
                stmt.setString(3, AccessDate);
                stmt.setString(4, ProcessTime);
                stmt.executeUpdate();
                */
            	
            	Cluster cluster = Cluster.builder().addContactPoint("127.0.0.1").build();
			    
			      //Creating Session object
			      Session session = cluster.connect("counterks");
			      System.out.println("connection made...");
			      
			      String query1 = "INSERT INTO Voterlist (VOTER, PARTY, AGE GROUP, BALLOT STATUS)"
			    			
         + " VALUES('?', '?', '?', '?');" ;
			      
			      session.setString(1, ClientAdd);
	                session.setString(2, Page);
	                session.setString(3, AccessDate);
	                session.setString(4, ProcessTime);
			      session.execute(query1);

                System.out.println("Data is inserted");
               
            } 
            //catch (ClassNotFoundException e)
           // {
            //    e.printStackTrace();
           // } 
            catch (Exception e) 
            {
                e.printStackTrace();
            } 



        }
    }



}