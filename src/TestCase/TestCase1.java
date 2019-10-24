package TestCase;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestCase1 {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File f=new File("C:\\Users\\Hp\\Desktop\\selenium.xlsx");
		FileInputStream fis=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(fis);
		Sheet s=wb.getSheetAt(0);
		int rowcount=s.getLastRowNum()-s.getFirstRowNum();
		for (int i=0;i<rowcount+1;i++)
		{
			Row row=s.getRow(i);
			for(int j=0;j<row.getLastCellNum();j++)
			System.out.println(row.getCell(j).getStringCellValue());
		}

	}

}
