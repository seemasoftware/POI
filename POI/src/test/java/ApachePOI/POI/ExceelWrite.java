package ApachePOI.POI;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceelWrite {

public void Write()throws IOException {
	File f = new File("../POI/Book2.xlsx");
	FileOutputStream fo = new FileOutputStream(f);
	XSSFWorkbook xs = new XSSFWorkbook();
	XSSFSheet xt = xs.createSheet("Test");
	
	for(int i=0;i<3;i=i+1)   // Loop for raw
	{
		XSSFRow xr = xt.createRow(i);
		for (int j=0; j<3;j=j+1)  // loop for column
		{
			XSSFCell xc =xr.createCell(j);
			xc.setCellValue("Write test");
		}
	}
	xs.write(fo);  // it will move data from workbook to output stream
	fo.flush(); // it will move data from output stream to file
	fo.close();  // for saving data
}


	public static void main(String[] args) throws IOException {
		ExceelWrite  e= new ExceelWrite(); 
		e.Write();
	}
}	
