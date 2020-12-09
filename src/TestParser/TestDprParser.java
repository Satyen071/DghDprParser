package TestParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestDprParser {
	public static class MyObject implements Comparable<MyObject> {

		  private Date dateTime;

		  public Date getDateTime() {
		    return dateTime;
		  }

		  public void setDateTime(Date datetime) {
		    this.dateTime = datetime;
		  }

		  @Override
		  public int compareTo(MyObject o) {
		    return getDateTime().compareTo(o.getDateTime());
		  }
		}

	public static void main(String[] args) throws FileNotFoundException, ParseException {
		// TODO Auto-generated method stub
		
		 FileInputStream fis;
		 XSSFWorkbook wb = null;
		try {
			fis = new FileInputStream(new File("/home/logicladder/Downloads/DPR_Dump_07-12-2020.xlsx"));
			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	      
	      XSSFSheet sheet = wb.getSheetAt(0);
	      
	      XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet1 = workbook.createSheet("DghDpr");
	      
	      FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
	      int rowCount=0;
	      
	      CellStyle cellStyle = sheet1.getWorkbook().createCellStyle();
	      Font font = sheet1.getWorkbook().createFont();
	      font.setBold(true);
//	      font.setFontHeightInPoints((short) 16);
	      cellStyle.setFont(font);
	      Row row = sheet1.createRow(0);
	      Cell cellTitle = row.createCell(0);
	   
	      cellTitle.setCellStyle(cellStyle);
	      cellTitle.setCellValue("Field Name");
	   
	      Cell cellOperator = row.createCell(1);
	      cellOperator.setCellStyle(cellStyle);
	      cellOperator.setCellValue("Operator Name");
	   
	      Cell cellActivity = row.createCell(2);
	      cellActivity.setCellStyle(cellStyle);
	      cellActivity.setCellValue("Activity Name");
	      
	      Cell cellDate = row.createCell(3);
	      cellDate.setCellStyle(cellStyle);
	      cellDate.setCellValue("Activity Date");
	      
	      Cell cellOilMt = row.createCell(4);
	      cellOilMt.setCellStyle(cellStyle);
	      cellOilMt.setCellValue("OIL_MT");
	      
	      Cell cellCondMt = row.createCell(5);
	      cellCondMt.setCellStyle(cellStyle);
	      cellCondMt.setCellValue("COND_MT");
	      
	      Cell cellOilBbl = row.createCell(6);
	      cellOilBbl.setCellStyle(cellStyle);
	      cellOilBbl.setCellValue("OIL_BBL");
	      
	      Cell cellCondBbl = row.createCell(7);
	      cellCondBbl.setCellStyle(cellStyle);
	      cellCondBbl.setCellValue("COND_BBL");
	      
	      Cell cellAssociatedGas = row.createCell(8);
	      cellAssociatedGas.setCellStyle(cellStyle);
	      cellAssociatedGas.setCellValue("ASSOCIATED_GAS_M3");
	      
	      Cell cellNonAssociatedGas = row.createCell(9);
	      cellNonAssociatedGas.setCellStyle(cellStyle);
	      cellNonAssociatedGas.setCellValue("NON_ASSOCIATED_GAS_M3");

System.out.println(sheet.getPhysicalNumberOfRows());

			List<MyObject> list = new ArrayList<MyObject>();
			SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yy");
			//String formattedDate = formatter.format();
			for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
				  Row row1 = sheet.getRow(i);
			//	  System.out.println(row1.getCell(3).getCellType());
				  if(row1.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY") &&
						  row1.getCell(0).getStringCellValue().equalsIgnoreCase("ALLORA" )) {
					  		try {
					  			MyObject obj = new MyObject();
					  			obj.setDateTime(formatter.parse(row1.getCell(3).getStringCellValue()));
								list.add(obj);
							} catch (ParseException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}  
				  }else {
					  continue;
				  }
				  
			}
			Collections.sort(list,Collections.reverseOrder());
	      for(int k=1;k<sheet.getPhysicalNumberOfRows();k++) {
//	    	  System.out.println(rw.getCell(0));
	    	  Row rw = sheet.getRow(k);
	    	  Row row1 = sheet1.createRow(++rowCount);
	    	  int columnCount = 0;
//	    	  rw.getCell(0);
//	    	  for(Cell cl:rw) {
//	    		Cell cell = row1.createCell(++columnCount);
//	    		if(rw.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY")) {
//	    		if(rw.getCell(0).getCellType() == CellType.STRING)
//	    		cell.setCellValue(rw.getCell(0).toString());
//	    		}else {
//	    			continue;
//	    		}
	    	  Calendar cal = Calendar.getInstance();
	    	  cal.setTime(list.get(0).getDateTime());
	    	  System.out.println(formatter.format(cal.getTime()));
	    	  String formatedDate = formatter.format(cal.getTime());
	    	  System.out.println(formatedDate);
	    	  if( formatedDate.equalsIgnoreCase(rw.getCell(3).getStringCellValue()) && rw.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY")) {
	    		 
	    		  for(int i=0;i<rw.getPhysicalNumberOfCells();i++) {
		    		  Cell cell = row1.createCell(i);
		    		  if(rw.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY")) {
		    			  if(rw.getCell(0).getCellType() == CellType.STRING ) {
		    				 
		    				  	cell.setCellValue(rw.getCell(i).toString());
		    				  
//		    					  cell.setCellValue(rw.getCell(i).toString());
		    				  
		    			  }
		    			  
		    		  }else {
		    			  continue;
		    		  }
		    		  
		    	  }
	    	  }
	    	  
	    	
////	    	  
	      }
	      
	      
    	  
    	  System.out.println(list.get(0).getDateTime());
    	  
    	  Collections.sort(list, new Comparator<MyObject>() {
    		  public int compare(MyObject o1, MyObject o2) {
    		      return o1.getDateTime().compareTo(o2.getDateTime());
    		  }
    		});
    	  System.out.println(list.get(0).getDateTime());
    	  System.out.println(sheet1.getPhysicalNumberOfRows());
	      
	      FileOutputStream out = new FileOutputStream("DghDpr.xlsx");
	      try {
			workbook.write(out);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
