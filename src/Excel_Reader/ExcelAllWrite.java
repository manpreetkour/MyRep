package Excel_Reader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelAllWrite {
	@Test
	public void allWrite() throws IOException{
		
		System.out.println("Writing into Excel files ..... ");
		String path=System.getProperty("user.dir")+"\\src\\XLS_Files\\TestData.xlsx";
		FileInputStream fis =new FileInputStream(path);
		XSSFWorkbook workbook= new XSSFWorkbook(fis);
		
		XSSFSheet sheet=workbook.getSheetAt(0);

		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			sheet.getRow(i).createCell(3).setCellValue("Selenium");
		}
		
		FileOutputStream fileOut = new FileOutputStream(path);
		workbook.write(fileOut);
	    fileOut.close();	
	}
}