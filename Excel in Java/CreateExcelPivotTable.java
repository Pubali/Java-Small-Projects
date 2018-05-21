package com.cassandra.practice.excelops;
import  com.cassandra.app.Extract_data;
import com.cassandra.practice.model.Voter;

import java.io.FileNotFoundException;
import java.io.IOException;

import java.io.*;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class CreateExcelPivotTable {

	
	public static void main(String[] args) throws IOException {
		
		//Create blank workbook
		
	      XSSFWorkbook workbook = new XSSFWorkbook(); 
	      //Create a blank sheet
	      XSSFSheet spreadsheet = workbook.createSheet( 
	      " Employee Info ");
	      Extract_data t = new Extract_data();
	      List<Voter> voterList = new ArrayList<Voter>();
	      voterList=t.getVoters(); 
	     // System.out.println("length"+voterList.size());
	      XSSFRow row=spreadsheet.createRow(1);
	      XSSFCell cell;
	      cell=row.createCell(1);
	      cell.setCellValue("voterid");
	      cell=row.createCell(2);
	      cell.setCellValue("party");
	      cell=row.createCell(3);
	      cell.setCellValue("age_group");
	      cell=row.createCell(4);
	      cell.setCellValue("ballot_status");
	     int i=2;
	      for(Voter voter :voterList){
	    	  
	    	  //System.out.println("Voterid"+voter.getVoterId());
	    	  XSSFRow datarow=spreadsheet.createRow(i);
		      XSSFCell datacell;
		      datacell=datarow.createCell(1);
		      datacell.setCellValue(voter.getVoterId());
		      datacell=datarow.createCell(2);
		      datacell.setCellValue(voter.getParty());
		      datacell=datarow.createCell(3);
		      datacell.setCellValue(voter.getAgeGroup());
		      datacell=datarow.createCell(4);
		      datacell.setCellValue(voter.getBallotStatus());
		     i++;
	    	  //create xssf row,add values to the row,
	      }
	      //print value on the console
	     
	     	      
	      //Write the workbook in file system
	      FileOutputStream out = new FileOutputStream( 
	      new File("C:/Users/pubali.bhaduri/Downloads/Writesheet.xls"));
	      workbook.write(out);
	      out.close();
	      System.out.println( 
	      "Writesheet.xlsx written successfully" );
	      Createpivot();
	   }

	public static  void  Createpivot() throws IOException
	{
		/* Read the input file that contains the data to pivot */
        FileInputStream input_document = new FileInputStream(new File("C:/Users/pubali.bhaduri/Downloads/Writesheet.xls"));    
        /* Create a POI HSSFWorkbook Object from the input file */
        XSSFWorkbook workbook = new XSSFWorkbook(input_document); 
        /* Read Data to be Pivoted - we have only one worksheet */
        XSSFSheet sheet = workbook.getSheetAt(0); 
        /* Get the reference for Pivot Data */
        AreaReference a=new AreaReference("B2:E9");
        /* Find out where the Pivot Table needs to be laced */
        CellReference b=new CellReference("I5");
        /* Create Pivot Table */
        XSSFPivotTable pivotTable = sheet.createPivotTable(a,b);
        /* Add filters */
        pivotTable.addReportFilter(0);
        pivotTable.addRowLabel(3);
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2); 
        /* Write Pivot Table to File */
        FileOutputStream output_file = new FileOutputStream(new File("C:/Users/pubali.bhaduri/Downloads/POI_XLS_Pivot_Example.xls")); 
        workbook.write(output_file);
        input_document.close(); 
	}

	}
