package com.sfit.doc;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.sfit.sheet.ServiceContent;

public class SheetReader {

	public static void getServiceContent(List<String> nameList, Sheet sheet){
		List<ServiceContent> serviceList = new ArrayList<ServiceContent>();
		int total = sheet.getLastRowNum() + 1;
		for (String name : nameList){
			for (int r=1; r<total; r++){
				Row rowTemp = sheet.getRow(r);
				if (rowTemp != null && rowTemp.getCell(0) != null){
					if (rowTemp.getCell(0).getStringCellValue() == name){
						ServiceContent serviceTemp = new ServiceContent();
						serviceTemp.setCompany(name);
						serviceTemp.setCatalog(rowTemp.getCell(1).getStringCellValue());
						//serviceTemp.setCategory(rowTemp.getCell(2).getStringCellValue());
						//serviceTemp.setDesc(rowTemp.getCell(3).getStringCellValue());
					}
				}
			}
		}
	}
	
	
	
}
