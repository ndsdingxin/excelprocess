package com.iflytek.excelprocess;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class WriteExcel {

	public static String tax[][] = { { "913401006679310226", "6mvi7z" }, { "91370211MA3NEU722T", "11ixk7" },
			{ "913401006679310227", "u2uxmy" } };

	static String str1 = "taxtoken";

	static String filenameString = "f:/taxtoken.xls";

	public static String writeexcel() {
		// 创建HSSFWorkbook对象(excel的文档对象)
		HSSFWorkbook wb = new HSSFWorkbook();
		// 建立新的sheet对象（excel的表单）
		HSSFSheet sheet = wb.createSheet(str1);
		// 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
		// 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
		 /*
	    tokenmsg (id integer primary key autoincrement,
	    taxnum varchar(50),
	    taxtoken varchar(50),
	    dizhi varchar(50),
	    remark varchar(50))
	     */
		HSSFRow row2 = sheet.createRow(0);
		// 创建单元格并设置单元格内容
		row2.createCell(0).setCellValue("id");
		row2.createCell(1).setCellValue("taxnum");
		row2.createCell(2).setCellValue("taxtoken");
		row2.createCell(3).setCellValue("dizhi");
		row2.createCell(4).setCellValue("remark");
		// 在sheet里创建第三行
		for (int i = 0; i < tax.length; i++) {
			HSSFRow row3 = sheet.createRow(i+1);
			row3.createCell(0).setCellValue(i+1);
			row3.createCell(1).setCellValue(tax[i][0]);
			row3.createCell(2).setCellValue(tax[i][1]);
			row3.createCell(3).setCellValue("无");
			row3.createCell(4).setCellValue("无");
			
		}
		// .....省略部分代码

		// 输出Excel文件
		try {
			FileOutputStream output = new FileOutputStream(filenameString);
			wb.write(output);
			output.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

}
