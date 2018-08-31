package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		File excel = new File("D:\\\\Testdata2018.xlsx");
		
		FileInputStream fis = new FileInputStream(excel);
		
        //File file = new File("D:\\Testdata2018.xlsx");
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		
		//XSSF
		
		

		//Workbook wb = WorkbookFactory.create(fis);
       XSSFSheet sheet = wb.getSheet("Sheet1");
        Row R = sheet.getRow(0);
        Cell C = R.getCell(0);
          //C.setCellValue("VARUN");
        //FileOutputStream os = new FileOutputStream(excel);
       // wb.write(os);
        String V = C.getStringCellValue();
        System.out.println(V);
        
        
        

		
		//Row r = s.getRow(0);
		
		//System.out.println(((Object) s.getMasterSheet()).getRow().getCellValue());
		
		
		
		

	}

}
