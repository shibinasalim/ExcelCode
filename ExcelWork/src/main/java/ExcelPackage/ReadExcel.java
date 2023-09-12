package ExcelPackage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	static FileInputStream fis;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String data=getStringData(0,0);
		   System.out.println(data);
				String data1=getStringData(0,1);
				System.out.println(data1);
				 int data4=getIntData(0,2);
				   System.out.println(4);
						int data5=getIntData(0,2);
						System.out.println(data5);
						 double data51=getDoubleData(0,3);
						   System.out.println(7);
								double data6=getDoubleData(0,3);
								System.out.println(data6);
				
	}
private static double getDoubleData(int e, int f) throws IOException {
		// TODO Auto-generated method stub
	fis=new FileInputStream("C:\\Users\\shibi\\OneDrive\\Desktop\\java\\shibina1.xlsx");
	wb=new XSSFWorkbook(fis);
	sheet=wb.getSheet("Sheet1");
	Row r=sheet.getRow(e);
	Cell g=r.getCell(f);
	return g.getNumericCellValue();
	}
public static int getIntData(int c, int d) throws  IOException{
	fis=new FileInputStream("C:\\Users\\shibi\\OneDrive\\Desktop\\java\\shibina1.xlsx");
	wb=new XSSFWorkbook(fis);
	sheet=wb.getSheet("Sheet1");
	Row r=sheet.getRow(c);
	Cell g=r.getCell(d);
	return (int)g.getNumericCellValue();
	
	}
public static String getStringData(int a ,int b) throws IOException
{
	fis=new FileInputStream("C:\\Users\\shibi\\OneDrive\\Desktop\\java\\shibina1.xlsx");
	
	wb=new XSSFWorkbook(fis);
	 sheet= wb.getSheet("Sheet1");
	Row r=sheet.getRow(a);
	Cell c=r.getCell(b);
	return c.getStringCellValue();
	

	
	}
}
