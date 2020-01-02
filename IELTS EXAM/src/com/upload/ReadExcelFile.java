package com.upload;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.json.simple.JSONArray;


import javafx.scene.control.Cell;

public class ReadExcelFile {
	
	public JSONArray readExcel(String file_path)
	{
		JSONArray array = new JSONArray();
		
		try {
			Workbook workbook = WorkbookFactory.create(new File(file_path));
			Sheet s = (Sheet) workbook.getSheetAt(0); // it reads first excel sheet
			XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
			
			DataFormatter dataFormatter = new DataFormatter();
			
			for(Row row:sheet)
			{
				JSONArray value = new JSONArray();
				for(org.apache.poi.ss.usermodel.Cell cell:row)
				{
					String cell_value = dataFormatter.formatCellValue(cell);
					value.add(cell_value);
				}
				
				array.add(value);
			}
			
			workbook.close();
			
			
			
			
		} catch (EncryptedDocumentException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		
		return array;
	}
}
