package ApachePOI.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	
	public void readExcel() throws IOException {
		File f = new File("../POI/Book1.xlsx");
		FileInputStream fi = new FileInputStream(f);
		XSSFWorkbook xs = new XSSFWorkbook(fi);
		XSSFSheet xt = xs.getSheetAt(0);
		int r = xt.getPhysicalNumberOfRows();
		
		for (int i=0;i<r;i=i+1)   // Loop for raw
		{
			XSSFRow xr = xt.getRow(i);
			int co = xr.getPhysicalNumberOfCells();
			for (int j=0; j<co;j=j+1)  // loop for column
			{
				XSSFCell xc =xr.getCell(j);
				System.out.println(xc.getStringCellValue());
			}
		}
	}
	
	public static void main(String[] args) throws IOException {
		ExcelRead e = new ExcelRead();
		e.readExcel();
	}

}
