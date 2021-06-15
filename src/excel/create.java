package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class create {
	
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//array for members supplied with water.
		 String[] line_1 = {"Joan","Kubia","Mark"};
		 String[] line_2 = {"Peter","kuria"};
		 String[] line_3 = {"malcom","zenia","tress","truex"};
		 String[] line_4 = {"zey"};
		 String[] line_5 = {"tre","yul"};
		 String[] line_6 = {"this","us"};
		 String[] lines[] = {line_1,line_2,line_3,line_4,line_5,line_6};
		 int x =lines.length;
		

		
		int z; int l;
		//create workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		
			//create fonts
		    XSSFFont defaultFont= wb.createFont();
		    defaultFont.setFontHeightInPoints((short)10);
		    defaultFont.setFontName("Arial");
		    defaultFont.setColor(IndexedColors.BLACK.getIndex());
		    defaultFont.setBold(false);
		    defaultFont.setItalic(false);

		    XSSFFont font= wb.createFont();
		    font.setFontHeightInPoints((short)10);
		    font.setFontName("Times New Roman");
		    font.setColor(IndexedColors.BLACK.getIndex());
		    font.setBold(true);
		    font.setItalic(false);
		    
		    XSSFCellStyle cs = wb.createCellStyle();
		    cs.setFont(font);
		    	
		    //create sheets
		for(l=0;l<x;l++) {
			XSSFSheet linesnames = wb.createSheet("line "+ (l+1));
			XSSFRow rh = linesnames.createRow(0);
			XSSFCell c1 =rh.createCell(0);	c1.setCellValue("Name");
			XSSFCell c2 = rh.createCell(1);	c2.setCellValue("Telephone no");
			c1.setCellStyle(cs);
			c2.setCellStyle(cs);
				System.out.print(lines[l].length + " ");
			  int y = lines[l].length;int r;
			 	for(r=0;r<y;r++){
			 		XSSFRow rn = linesnames.createRow(r+1);
			  rn.createCell(0).setCellValue(lines[l][r]);	
			 		}
				
			}
		File dt = new File("data.xlsx");
		try {
			
			FileOutputStream fs = new FileOutputStream(dt);
			wb.write(fs);
			fs.close();
			wb.close();
			System.out.print(lines[0][2]);
			System.out.print("sana");
			
		}catch(IOException e) {
			e.printStackTrace();}
			/*try {
				FileInputStream fi = new FileInputStream(dt);
			}catch(IOException z1) {
				z1.printStackTrace();
			}*/
			//create files for each line
			File bills1 = new File("BILLS\\Line_1.xlsx");
			File bills2 = new File("BILLS\\Line_2.xlsx");
			File bills3 = new File("BILLS\\Line_3.xlsx");
			File bills4 = new File("BILLS\\Line_4.xlsx");
			File bills5 = new File("BILLS\\Line_5.xlsx");
			File bills6 = new File("BILLS\\Line_6.xlsx");

	

}
	}

