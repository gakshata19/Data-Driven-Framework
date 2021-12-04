package Data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public FileInputStream fis;
	public XSSFWorkbook wb;
	public XSSFSheet sheet;
	public ExcelReader() {
		File f=new File("");
		try {
			fis=new FileInputStream(f);
			wb=new XSSFWorkbook(fis);			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public String getData(int sheetindex, int row, int col) {
		sheet=wb.getSheetAt(sheetindex);
		String data=sheet.getRow(row).getCell(col).getStringCellValue();
		return data;
	}
	
	public int rowcount(int sheetindex) {
		int row=wb.getSheetAt(sheetindex).getLastRowNum();
		row=row+1;
		return row;
	}
	
	public int colcount(int sheetindex) {
		int col=wb.getSheetAt(sheetindex).getRow(0).getLastCellNum();
		col=col+1;
		return col;
	}
}
