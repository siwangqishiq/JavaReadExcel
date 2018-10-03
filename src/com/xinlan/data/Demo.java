package com.xinlan.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;

public class Demo {
	public static void main(String args[]) throws Exception {
		File file = new File("º¼ÖÝ.xls");

		if (file.exists())
			System.out.println(file.length());

		InputStream fis = new FileInputStream(file);

		HSSFWorkbook workbook = new HSSFWorkbook(fis);

		// for(int i = 0 ; i< workbook.getNumberOfSheets() ; i++){
		HSSFSheet sheet = workbook.getSheetAt(0);

		if (sheet == null)
			return;
		// System.out.println("xssfSheet.getLastRowNum() = "
		// +sheet.getLastRowNum());
		List<Sexy> list = new ArrayList<Sexy>();
		for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
			HSSFRow xssfRow = sheet.getRow(rowNum);
			if (xssfRow == null)
				continue;

			Sexy item = new Sexy();
			item.setId(rowNum);

			HSSFCell cell1 = xssfRow.getCell(1);
			if (cell1 == null)
				continue;
			System.out.println(cell1.getStringCellValue());
			item.setTitle(cell1.getStringCellValue());

			HSSFCell cell2 = xssfRow.getCell(2);
			if (cell2 == null)
				continue;
			System.out.println(cell2.getStringCellValue());
			item.setType(cell2.getStringCellValue());
			
			HSSFCell cell3 = xssfRow.getCell(3);
			if (cell3 != null){
				System.out.println(cell3.getStringCellValue());
				item.setArea(cell3.getStringCellValue());
			}
			
			HSSFCell cell4 = xssfRow.getCell(4);
			if (cell4 != null){
				System.out.println(cell4.getStringCellValue());
				item.setLocation(cell4.getStringCellValue());
			}
			
			HSSFCell cell5 = xssfRow.getCell(5);
			if (cell5 != null){
				System.out.println(cell5.getStringCellValue());
				item.setSrc(cell5.getStringCellValue());
			}
			
			HSSFCell cell6 = xssfRow.getCell(6);
			if (cell6 != null){
				System.out.println(cell6.getStringCellValue());
				item.setNum(cell6.getStringCellValue());
			}
			
			HSSFCell cell7 = xssfRow.getCell(7);
			if (cell7 != null){
				System.out.println(cell7.getStringCellValue());
				item.setAge(cell7.getStringCellValue());
			}
			
			HSSFCell cell8 = xssfRow.getCell(8);
			if (cell8 != null){
				System.out.println(cell8.getStringCellValue());
				item.setStuff(cell8.getStringCellValue());
			}
			
			HSSFCell cell9 = xssfRow.getCell(9);
			if (cell9 != null){
				System.out.println(cell9.getStringCellValue());
				item.setFace(cell9.getStringCellValue());
			}
			
			HSSFCell cell10 = xssfRow.getCell(10);
			if (cell10 != null){
				System.out.println(cell10.getStringCellValue());
				item.setService(cell10.getStringCellValue());
			}
			
			HSSFCell cell11 = xssfRow.getCell(11);
			if (cell11 != null){
				System.out.println(cell11.getStringCellValue());
				item.setPrice(cell11.getStringCellValue());
			}
			
			HSSFCell cell12 = xssfRow.getCell(12);
			if (cell12 != null){
				System.out.println(cell12.getStringCellValue());
				item.setTime(cell12.getStringCellValue());
			}
			
			HSSFCell cell13 = xssfRow.getCell(13);
			if (cell13 != null){
				System.out.println(cell13.getStringCellValue());
				item.setEnv(cell13.getStringCellValue());
			}
			
			HSSFCell cell14 = xssfRow.getCell(14);
			if (cell14 != null){
				System.out.println(cell14.getStringCellValue());
				item.setSafty(cell14.getStringCellValue());
			}
			
			HSSFCell cell15 = xssfRow.getCell(15);
			if (cell15 != null){
				System.out.println(cell15.getStringCellValue());
				item.setMobile(cell15.getStringCellValue());
			}
			
			HSSFCell cell16 = xssfRow.getCell(16);
			if (cell16 != null){
				System.out.println(cell16.getStringCellValue());
				item.setScore(cell16.getStringCellValue());
			}
			
			HSSFCell cell17 = xssfRow.getCell(17);
			if (cell17 != null){
				System.out.println(cell17.getStringCellValue());
				item.setDesc(cell17.getStringCellValue());
			}

			list.add(item);
		}

		System.out.println("total = " + list.size());
		
		for(int i = 0 ;i < list.size() ; i++){
			System.out.println(JSON.toJSONString(list.get(i)));
		}
		
//		String saveString = JSONArray.toJSONString(list);
//		System.out.println(saveString);
		
		
//		FileOutputStream fos = new FileOutputStream(new File("data.json"));
//		fos.write(saveString.getBytes("UTF-8"));
//		fos.close();
		// }
		fis.close();
	}

}// end class
