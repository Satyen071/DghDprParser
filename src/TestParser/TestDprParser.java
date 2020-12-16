package TestParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
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
import java.util.Properties;

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

	public static List<MyObject> getFieldLatestData(XSSFSheet sheet, SimpleDateFormat formatter,String fieldName) {
		List<MyObject> list = new ArrayList<MyObject>();

		for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
			  Row row1 = sheet.getRow(i);
		//	  System.out.println(row1.getCell(3).getCellType());
			  if(row1.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY") && 
					 row1.getCell(0).getStringCellValue().equalsIgnoreCase(fieldName )) {
				  		try {
				  			MyObject obj = new MyObject();
				  			obj.setDateTime(formatter.parse(row1.getCell(3).getStringCellValue()));
							list.add(obj);
						} catch (ParseException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}  
					 }
			  
			  
		}
		  Collections.sort(list,Collections.reverseOrder());
		  return list;
	}
	
	
	public static XSSFSheet removeEmptyRows(XSSFSheet sheet) {
	    Boolean isRowEmpty = Boolean.FALSE;
	    for(int i = 0; i <= sheet.getLastRowNum(); i++){
	      if(sheet.getRow(i)==null || sheet.getRow(i).getLastCellNum()==-1){
	        isRowEmpty=true;
	        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
	        i--;
	        continue;
	      }
//	      for(int j =0; j<sheet.getRow(i).getLastCellNum();j++){
//	        if(sheet.getRow(i).getCell(j).getCellType() == CellType.BLANK || 
//	        sheet.getRow(i).getCell(j).toString().trim().equals("")){
//	          isRowEmpty=true;
//	        }else {
//	          isRowEmpty=false;
//	          break;
//	        }
//	      }
//	      if(isRowEmpty==true){
//	        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
//	        i--;
//	      }
	    }
	    return sheet;
	  }
	
	public static void main(String[] args) throws FileNotFoundException, ParseException {
		// TODO Auto-generated method stub
		
		 FileInputStream fis;
		 XSSFWorkbook wb = null;
		try {
			fis = new FileInputStream(new File("/home/logicladder/Downloads/DPRDump15Dec.xlsx"));
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
//			List<MyObject> list = new ArrayList<MyObject>();
			SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yy");
			String[] fieldName = {"AAP-ON-94/1","ALLORA","ASJOL","BAKROL","BAOLA","BHANDUT","BOKARO","CAMBAY","CB-ON/2","CB-ON/3", "CB-ON/7", 
					"CB-ONN-2000/1","CB-ONN-2001/1","CB-ONN-2002/3","CB-ONN-2003/1" 
					,"CB-ONN-2003/2","CB-ONN-2004/1","CB-ONN-2004/2",
					"CB-ONN-2004/3","CB-ONN-2005/9","CB-OS/2","CB-OSN-2003/1",
					"CY-ONN-2002/2","DHOLASAN","DHOLKA","HAZIRA","JHARIA","KANAWARA","KARJISAN","KG-DWN-98/2",
					"KG-ONN-2003/1","KG-OSN-2001/3","KHARSANG","LOHAR","MODHERA", 
					"N.BALOL","NORTH KATHANA","OGNAJ","PY-1","RANIGANJ EAST","RANIGUNJ SOUTH","RAVVA","RJ-ON/6","RJ-ON-90/1",
					"SOHAGPUR EAST","SOHAGPUR WEST","UNAWA","WAVEL" };
			List<MyObject> list = new ArrayList<TestDprParser.MyObject>();
			for(String str:fieldName) {
				
			
			 list = getFieldLatestData(sheet,formatter,str);
			System.out.println(list.get(0).getDateTime());
	      for(int k=1;k<sheet.getPhysicalNumberOfRows();k++) {
	    	  Row rw = sheet.getRow(k);
	    	  Calendar cal = Calendar.getInstance();
	    	  cal.setTime(list.get(0).getDateTime());
	    	  String formatedDate = formatter.format(cal.getTime());
	    	  if( formatedDate.equalsIgnoreCase(rw.getCell(3).getStringCellValue()) 
	    			  && rw.getCell(2).getStringCellValue().equalsIgnoreCase("ACTUAL PRODUCTION FOR THE DAY")
	    			  && rw.getCell(0).getStringCellValue().equalsIgnoreCase(str.trim() )) {
	    		  		System.out.println(str);
	    			  System.out.println(rw.getCell(3).getStringCellValue());
	    			  Row row1 = sheet1.createRow(++rowCount);
	    			  	for(int i=0;i<rw.getPhysicalNumberOfCells();i++) {
	    			  		
		    		  			Cell cell = row1.createCell(i);
		    				  	cell.setCellValue(rw.getCell(i).toString());
		    			  
	    			  	}
		    		  
	    	  }
	      }
	}
	      
	      
    	  
//    	  System.out.println(list.get(0).getDateTime());
    	  
//    	  Collections.sort(list, new Comparator<MyObject>() {
//    		  public int compare(MyObject o1, MyObject o2) {
//    		      return o1.getDateTime().compareTo(o2.getDateTime());
//    		  }
//    		});
//    	  System.out.println(list.get(0).getDateTime());
    	  System.out.println(sheet1.getPhysicalNumberOfRows());
//    	  XSSFWorkbook workbook1 = new XSSFWorkbook();
//	        XSSFSheet sheet2 = workbook1.createSheet("DghDpr1");
    	   TestDprParser.removeEmptyRows(sheet1);
	      
	      FileOutputStream out = new FileOutputStream("DghDpr.xlsx");
	      try {
			workbook.write(out);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
