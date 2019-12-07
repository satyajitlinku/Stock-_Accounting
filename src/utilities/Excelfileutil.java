package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelfileutil{
	Workbook wb;
	public  Excelfileutil()throws Exception{
	FileInputStream fi=new FileInputStream("C:\\Users\\satyajit.k\\Documents\\Satyajittdata.xlsx");
	 wb=new XSSFWorkbook(fi);
	}

public int rowcount(String sheetname){
	return wb.getSheet(sheetname).getLastRowNum();
}
public int colcount(String sheetname){
	return wb.getSheet(sheetname).getRow(0).getLastCellNum();
}
public String getData(String sheetname,int row,int column){
	String data;
	if (wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC) {
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
		data=String.valueOf(celldata);
		
	} else {
		data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();

	}
	return data;
	}
public void setData(String sheetname,int row,int column,String status) throws Exception  {
	Sheet sh= wb.getSheet(sheetname);
	Row rownum= sh.getRow(row);
	Cell cell=rownum.createCell(column);
	cell.setCellValue(status);

	if (status.equalsIgnoreCase("pass"))
	{
		//Create a cell Style
		CellStyle style=wb.createCellStyle();
		//create a font
		Font font=wb.createFont();
		//apply color to the text
		font.setColor(IndexedColors.GREEN.getIndex());
		//apply bold to the text
		font.setBold(true);
		//set font
		style.setFont(font);
		// set cell style
		rownum.getCell(column).setCellStyle(style);
	}else
		if(status.equalsIgnoreCase("Fail"))
		{
			//Create a cell Style
			CellStyle style=wb.createCellStyle();
			//create a font
			Font font=wb.createFont();
			//apply color to the text
			font.setColor(IndexedColors.RED.getIndex());
			//app bold to the text
			font.setBold(true);
			//set font
			style.setFont(font);
			// set cell style
			rownum.getCell(column).setCellStyle(style);
		}else
			if(status.equalsIgnoreCase("Not Executed"))
			{
				//Create a cell Style
				CellStyle style=wb.createCellStyle();
				//create a font
				Font font=wb.createFont();
				//apply color to the text
				font.setColor(IndexedColors.BLUE.getIndex());
				//app bold to the text
				font.setBold(true);
				//set font
				style.setFont(font);
				// set cell style
				rownum.getCell(column).setCellStyle(style);
			}
	
	FileOutputStream fos=new FileOutputStream("C:\\Users\\satyajit.k\\Documents\\Satyajittdata.xlsx");
	wb.write(fos);
	fos.close();
}

}


































	

	
	
	
	
	
	
	
	
	
	
	


