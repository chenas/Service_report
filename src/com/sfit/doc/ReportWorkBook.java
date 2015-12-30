package com.sfit.doc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sfit.sheet.Summary;

public class ReportWorkBook {

	private static String[] sheet_service_names = {"机房服务", "机柜服务", "CTP软件服务", "CTPmini"};
	private static String[] sheet_list_names = {"机房服务名单", "机柜服务名单", "CTP软件服务名单", "CTPmini名单"};
	private static String[] sheet_other_names = {"客户名单", "概述"};
	
	private int type = 1;
	//年报的年份
	private int curYear;
	//年报下个年份
	private int nextYear;
	//开始时间
	private String beginDate;
	//结束时间
	private String endDate;
	
	//需要生成报告的公司列表
	private List<String> companyList;
	
	//概述
	private List<Summary> summaryList;
	
	private List<String> sheetNamesService;
	private List<String> sheetNamesList;
	private List<String> sheetNamesOthere;
		
	private Sheet sheetSummary;
	
	private Map<List<String>, Sheet> sheetService;
	
	/**
	 * 
	 * @param workBookPath
	 * @param type 季报：1  	年报： 2
	 */
	public ReportWorkBook(String workBookPath, int type, String beginDate, String endDate, int year){
		this.curYear = year;
		this.nextYear = year + 1;
		this.beginDate = beginDate;
		this.endDate = endDate;
		sheetService = new HashMap<List<String>, Sheet>();
		InputStream inp = null;
		try {
			inp = new FileInputStream(workBookPath);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
			System.out.println(e1.toString());
		}
		Workbook wb;
		try {
			wb = new XSSFWorkbook(inp);
			initSheet(wb);
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println(e.toString());
		}
//		sheetNamesService = Arrays.asList(sheet_service_names);
//		sheetNamesOthere = Arrays.asList(sheet_other_names);
//		sheetNamesList = Arrays.asList(sheet_list_names);
	}
	
	private void initSheet(Workbook wb){
		companyList = getNameList(wb.getSheet(sheet_other_names[0]));
		//年终概述
		if (type == 2){
			sheetSummary = wb.getSheet(sheet_other_names[1]);
			summaryList = getSummary(sheetSummary);
		}
				
		for (int i=0; i<sheet_list_names.length; i++){
			Sheet sheetTemp = wb.getSheet(sheet_list_names[i]);
			if (sheetTemp == null)
				continue;
			List<String> nameList = new ArrayList<String>();
			//将名单列表加入list
			nameList = getNameList(sheetTemp);
			//服务内容表格
			sheetTemp = wb.getSheet(sheet_service_names[i]);
			sheetService.put(nameList, sheetTemp);
		}
	}

	/**
	 * 遍历名单表格
	 * 只有一列
	 * @param sheet
	 * @return
	 */
	private List<String> getNameList(Sheet sheet){
		List<String> nameList  = new ArrayList<String>();
		int totalRow = sheet.getLastRowNum() + 1;
		//将名单列表加入list
		for (int r=0; r<totalRow; r++){
			if(sheet.getRow(r) != null 
					&& sheet.getRow(r).getCell(0)!= null
					&& sheet.getRow(r).getCell(0).getStringCellValue() != null)
			{
				nameList.add(sheet.getRow(r).getCell(0).getStringCellValue());
			}
		}
		return nameList;
	}
	
	/**
	 * 遍历年报概述表格
	 * @param sheet
	 * @return
	 */
	private List<Summary> getSummary(Sheet sheet){
		List<Summary> summaryList = new ArrayList<Summary>();
		for (int r=0; r<sheet.getLastRowNum()+1; r++){
			Row rowTemp = sheet.getRow(r);
			if (rowTemp != null && rowTemp.getCell(0) != null){
				Summary summaryTemp = new Summary();
				summaryTemp.setCompany(rowTemp.getCell(0).getStringCellValue());
				summaryTemp.setSummary(rowTemp.getCell(2).getStringCellValue());
				summaryTemp.getTITLE_CONTENT_11().replaceAll("yyyy", String.valueOf(curYear)).replace("nnnn", String.valueOf(nextYear));
			}
		}		
		return summaryList;
	}
	
	public List<String> getSheetNamesService() {
		return sheetNamesService;
	}

	public void setSheetNamesService(List<String> sheetNamesService) {
		this.sheetNamesService = sheetNamesService;
	}

	public List<String> getSheetNamesList() {
		return sheetNamesList;
	}

	public void setSheetNamesList(List<String> sheetNamesList) {
		this.sheetNamesList = sheetNamesList;
	}

	public List<String> getSheetNamesOthere() {
		return sheetNamesOthere;
	}

	public void setSheetNamesOthere(List<String> sheetNamesOthere) {
		this.sheetNamesOthere = sheetNamesOthere;
	}

	public Sheet getSheetSummary() {
		return sheetSummary;
	}

	public void setSheetSummary(Sheet sheetSummary) {
		this.sheetSummary = sheetSummary;
	}

	public Map<List<String>, Sheet> getSheetService() {
		return sheetService;
	}

	public void setSheetService(Map<List<String>, Sheet> sheetService) {
		this.sheetService = sheetService;
	}

	public int getType() {
		return type;
	}

	public void setType(int type) {
		this.type = type;
	}

	public int getCurYear() {
		return curYear;
	}

	public void setCurYear(int curYear) {
		this.curYear = curYear;
	}

	public int getNextYear() {
		return nextYear;
	}

	public void setNextYear(int nextYear) {
		this.nextYear = nextYear;
	}

	public String getBeginDate() {
		return beginDate;
	}

	public void setBeginDate(String beginDate) {
		this.beginDate = beginDate;
	}

	public String getEndDate() {
		return endDate;
	}

	public void setEndDate(String endDate) {
		this.endDate = endDate;
	}

	public List<String> getCompanyList() {
		return companyList;
	}

	public void setCompanyList(List<String> companyList) {
		this.companyList = companyList;
	}

	public List<Summary> getSummaryList() {
		return summaryList;
	}

	public void setSummaryList(List<Summary> summaryList) {
		this.summaryList = summaryList;
	}
	
}

