package Excel_Reader;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	public static void main(String[] args) throws IOException {
	    String fName=System.getProperty("user.dir")+"\\src\\XLS_Files\\TestData.xlsx";
		XSSFWorkbook wb=new XSSFWorkbook(fName);
		XSSFSheet sheet=wb.getSheetAt(0);
		
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			System.out.println(sheet.getRow(i).getCell(0)+"  --  "+sheet.getRow(i).getCell(1));
		}
	}
}