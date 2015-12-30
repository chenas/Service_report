package com.sfit.doc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


public class GenerateRepoter {

	private String custome_name_temp;
	//private static String[] sheet_names = {"客户名单","打包名单","CTP软件服务名单","机房服务名单","机柜服务名单","机柜服务","机房服务","CTP软件服务", "概述"};
	private static String[] sheet_names = {"客户名单","打包名单","CTP软件服务名单","机房服务名单","机柜服务名单","机柜服务","机房服务","CTP软件服务","CTPmini名单","CTPmini", "概述"};
	// above is standard order of sheet
	private int current_year;
	private int next_year;
	private boolean soft_service = false;
	private boolean machinery_service = false;
	private boolean cabinet_service = false;
	private String service_list;
	private String summary_desc;
	//above is for 概述
	
	private Sheet custome_list; //客户名单
	private Sheet soft_list; //CTP软件服务名单
	private Sheet machinery_list; //机房服务名单
	private Sheet cabinet_list; //机柜服务名单
	private Sheet soft_desc; //CTP软件租赁
	private Sheet machinery_desc; //机房资源租赁
	private Sheet cabinet_desc; //机柜资源租赁
	private Sheet summary; //概述
	
	//add by chen.guoshu
	private Sheet ctpmini_list; //ctpmini名单
	private Sheet ctpmini_desc; //ctpmini服务
	private boolean ctpmini_service = false; //是否有ctpmini服务
	
	private String[] styles;
	private String fileInfolder = "";
	private String modal = "";
	private String sDate;
	private String eDate;
	private String season = "";
public void GenerateReport(int reportYear, String filepath, String fileInFolder, String modalPath, String startDate, String endDate, String season){
		String file = filepath;
		this.fileInfolder = fileInFolder;
		this.modal = modalPath;
		this.sDate = startDate;
		this.eDate = endDate;
		this.season = season;
		//current_year = Calendar.getInstance().get(Calendar.YEAR);
		current_year = reportYear;
		next_year = current_year+1;
		
		try{			
			//InputStream inp = new FileInputStream("F:/报告生成程序/对外服务报告生成数据2013年报参考——变化.xlsx");
			InputStream inp = new FileInputStream(file);
			//Workbook wb = WorkbookFactory.create(inp);
			Workbook wb = new XSSFWorkbook(inp);
			
			int sheet_number = wb.getNumberOfSheets();					
			for(int i=0;i<sheet_number;i++){ // give sheets values
				String name_temp = wb.getSheetName(i);
				int index = findIndex(name_temp);
				switch(index){
				case 0: custome_list = wb.getSheetAt(i);break;
				case 1: wb.getSheetAt(i);break;
				case 2: soft_list = wb.getSheetAt(i);break;
				case 3: machinery_list = wb.getSheetAt(i);break;
				case 4: cabinet_list = wb.getSheetAt(i);break;
				case 5: cabinet_desc = wb.getSheetAt(i);break;
				case 6: machinery_desc = wb.getSheetAt(i);break;
				case 7: soft_desc = wb.getSheetAt(i);break;
				case 8: ctpmini_list = wb.getSheetAt(i); break;
				case 9: ctpmini_desc = wb.getSheetAt(i); break;
				
				case 10: summary = wb.getSheetAt(i);break;
				}
			}
			
			genarateWord();
						
		}catch(Exception e){
			e.printStackTrace();
		}	
	}
	
//生成word
private void genarateWord() {
		
	int custome_number = custome_list.getLastRowNum()+1;
	//int custome_number = 5;
	for(int i=0;i<custome_number;i++){
		custome_name_temp = custome_list.getRow(i).getCell(0).getStringCellValue();// get custome names
		combineService_list(custome_name_temp); // give service_list a value
		//System.out.println(custome_name_temp);
		try{
			// 生成文件夹
			
			File f = new File(fileInfolder+"/生成文件");
			if(!f.exists()){
				if(f.mkdir()){
					//System.out.println("Genarate A New Folder Succeed!");
				}
			}
			FileOutputStream os = null;
			if (season == "" || "".equals(season)){
				//年报
				os = new FileOutputStream(fileInfolder+"/生成文件/对外服务年报报告("+custome_name_temp+").docx");// create the word
			}
			else 
			{
				//季报
				os = new FileOutputStream(fileInfolder+"/生成文件/"+current_year+"年第"+season+"季度对外服务报告("+custome_name_temp+").docx");// create the word
			}
			
			
			InputStream inp = new FileInputStream(modal);			
			//InputStream inp = new FileInputStream("F:/报告生成程序/对外服务年报模板（内部版）.docx");
			
			XWPFDocument doc = new XWPFDocument(inp);
			
			// get styles
			styles = new String[6];
			List<XWPFParagraph> paras = doc.getParagraphs();
			
			for(int j=1;j<7;j++){
				XWPFParagraph xp1;
				xp1 = paras.get(paras.size()-j) ;
				//System.out.println(xp1.getParagraphText());
				//System.out.println(xp1.getStyle());
				styles[j-1] = xp1.getStyle();				
			}						
			//System.out.println(Arrays.toString(styles));
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			doc.removeBodyElement(doc.getBodyElements().size()-1);
			
			// Head 1. 概述
			addSummary(doc);
			
			doc.write(os);
			os.flush();
			os.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}
}

private void combineService_list(String custome_name_temp) {
	
	soft_service = false;
	machinery_service = false;
	cabinet_service = false;
	ctpmini_service = false;
	
	// CTP soft
	int soft_number = soft_list.getLastRowNum()+1;
	for(int i=0;i<soft_number;i++){
		if(custome_name_temp.equals(soft_list.getRow(i).getCell(0).getStringCellValue())){
			soft_service = true;
			break;
		}
	}
	// 机房
	//int machinery_number = machinery_list.getLastRowNum()+1;
	int machinery_number = machinery_list.getLastRowNum();
	if(custome_name_temp.equals("财达期货")){
		System.out.println(machinery_service);
		System.out.println(machinery_number);
	}
	for(int j=0;j<machinery_number;j++){
		System.out.println(machinery_list.getRow(j).getCell(0).getStringCellValue());
		if(custome_name_temp.equals(machinery_list.getRow(j).getCell(0).getStringCellValue())){
			if(custome_name_temp.equals("财达期货")){
				System.out.println("xx"+machinery_list.getRow(j).getCell(0).getStringCellValue()+"xx");
				System.out.println(j);
			}
			machinery_service = true;
			break;
		}
	}
	// 机柜
	int cabinet_number = cabinet_list.getLastRowNum()+1;
	for(int k=0;k<cabinet_number;k++){
		if(custome_name_temp.equals(cabinet_list.getRow(k).getCell(0).getStringCellValue())){
			cabinet_service = true;
			break;
		}
	}
	
	//added by chen.guoshu
	//ctpmini 
	int ctpmini_number = ctpmini_list.getLastRowNum()+1;
	for(int t=0; t<ctpmini_number; t++){
		if(ctpmini_list.getRow(t) != null && ctpmini_list.getRow(t).getCell(0)!=null)
		{
			if(custome_name_temp.equals(ctpmini_list.getRow(t).getCell(0).getStringCellValue())){
				ctpmini_service = true;
				break;
			}
		}
	}
	
	service_list = "";
	service_list += soft_service == true? "CTP软件服务,":"";
	service_list += machinery_service == true? "机房服务,":"";
	service_list += cabinet_service == true? "机柜服务,":"";	
	
	service_list += ctpmini_service == true? "CTP MINI服务,":"";	
	
	//added by chen.guoshu
	//service_list += 
//	System.out.println(soft_service);
//	System.out.println(machinery_service);
//	System.out.println(cabinet_service);
	
	//季报不执行，年报才执行。
	// get summary index
	if (season == "" || "".equals(season)){
		if(summary.getLastRowNum()>0){
			int summary_number = summary.getLastRowNum()+1;
			for(int i=1;i<summary_number;i++){
				System.out.println(custome_name_temp);
				if(null != summary.getRow(i) && null != summary.getRow(i).getCell(0))
				{
					System.out.println(summary.getRow(i).getCell(0).getStringCellValue());
					if(custome_name_temp.equals(summary.getRow(i).getCell(0).getStringCellValue())){
						summary_desc = summary.getRow(i).getCell(2).getStringCellValue();
						break;
					}
				}
			}
		}
	}
	//System.out.println(Arrays.toString(summary_desc.split("\n")));
	//System.out.println(summary_desc);
}

//写标题
private void addSummary(XWPFDocument doc) {
	//季报不执行，年报执行
	// 1.概述 +  内容

	if (season == "" || "".equals(season)){
	addHeader(doc, "概述");
	//addEnter(doc);
	String body_summary_1 = "尊敬的客户，感谢您使用上海期货信息技术有限公司提供的服务，本公司秉承“预防为主、夯实基础、稳健运行、稳步发展”的方针竭诚为您服务！";
	//addHeaderContent(doc, body_summary_1);
	addContent(doc, body_summary_1);
	
//	// 1.1 总体运行情况描述
	addSubHeader(doc, "总体运行情况描述");
	String body_summary_11 = "尊敬的客户，您在"+current_year+"年度所采购的我司的服务项分别为"+service_list+"这些服务"+current_year+"年度的总体运行情况均达到了相应合同条款的承诺，感谢您在"+current_year+"年度对于我司工作的配合，希望您在"+next_year+"年度继续与我司精诚合作，为相关服务的稳定可靠运行共同努力！";
	addContent(doc, body_summary_11);
	addEnter(doc);
	
	// 1.2 满意度调查结果分析
	addSubHeader(doc, "满意度调查结果分析");
	String[] summary_desc_array = summary_desc.split("\n");
	for(int i=0;i<summary_desc_array.length;i++){
		addContent(doc, summary_desc_array[i].trim());
	}	
	}
	
	// 2. 机柜服务
	addHeader(doc, "机柜服务");
	if(!cabinet_service){
		addBoldContent(doc, "贵公司未采购该服务。");
		addEnter(doc);
	}else{
		//System.out.println("x1");
		listService(cabinet_desc,doc);		
	}
	// 3. 机房服务
	addHeader(doc, "机房服务");	
	if(!machinery_service){
		addBoldContent(doc, "贵公司未采购该服务。");
		addEnter(doc);
	}else{
		//System.out.println("x12");
		listService(machinery_desc, doc);
		
		//addBoldContent(doc, "机房服务更多相关服务信息可参照对外服务报告附件《机房/机柜服务客户对外服务报告（所有客户）》。");
		if (season == "" || "".equals(season))
		{
			addBoldContent(doc, "机柜服务更多相关服务信息可参照对外服务报告附件《机房/机柜服务客户对外服务报告（所有客户）-"+current_year+"年度》。");
		}
		else
		{
			addBoldContent(doc, "机柜服务更多相关服务信息可参照对外服务报告附件《资源租赁客户对外服务报告（所有客户）-"+current_year+"年"+season+"》");
		}
		addEnter(doc);
	}
	// 4. 软件服务
	addHeader(doc, "CTP软件服务");
	if(!soft_service){
		addBoldContent(doc, "贵公司未采购该服务。");
		addEnter(doc);
	}else{
		//System.out.println("x123");
		listService(soft_desc, doc);		
	}
	
	//added by chen.guoshu
	//CTP MINI服务
	addHeader(doc, "CTP MINI服务");
	if(!ctpmini_service){
		addBoldContent(doc, "贵公司未采购该服务。");
	}else{
		listService(ctpmini_desc, doc);
	}
	
	
	addEnding(doc);
}

//写入服务内容
private void listService(Sheet sheet,XWPFDocument doc) {
	//System.out.println(sheet.getLastRowNum());
	int cabinet_number = sheet.getLastRowNum()+1;			
	List<String> categories = new ArrayList<String>();
	for(int i=1;i<cabinet_number;i++){
		//System.out.println(sheet.getRow(i).getCell(2).getStringCellValue());
		//System.out.println(sheet.getRow(i).getCell(2));
		if(sheet.getRow(i) != null && sheet.getRow(i).getCell(2) != null ){
			String temp = sheet.getRow(i).getCell(2).getStringCellValue();
			//System.out.println(temp);
			if(!categories.contains(temp)){
				categories.add(temp);
			}
		}
	}
	int tempFlag = 0;
	//int count = categories.size();
	for(Iterator<String> it=categories.iterator();it.hasNext();){
		String t = it.next();					
		if(!calculateTime(doc,t,sheet)){ // 计算该公司共使用了多少次该类服务 + 并列出来			
			continue;
		}
		tempFlag++;
		addEnter(doc);		
	}
	if(tempFlag == 0){
		if (season == "" || "".equals(season))
		{
			addBoldContent(doc, "贵公司本年度未有服务记录，感谢贵公司对我司该项服务的一贯支持和信任。");
		}
		else
		{
			addBoldContent(doc, "贵公司本季度未有服务记录，感谢贵公司对我司该项服务的一贯支持和信任。");
		}
		
		addEnter(doc);
	}
}

//统计服务次数，并且写下服务内容
private boolean calculateTime(XWPFDocument doc, String t, Sheet sheet) {
	int time = 0;
	List<Integer> indexs = new ArrayList<Integer>();
	int cabinet_number = sheet.getLastRowNum()+1;
	//System.out.println(cabinet_number);
	for(int i=1;i<cabinet_number;i++){
		if (sheet.getRow(i) == null || sheet.getRow(i).getCell(2) == null)
			continue;
		String temp = sheet.getRow(i).getCell(2).getStringCellValue();
		String cpy_name = "";
		try{
			cpy_name = sheet.getRow(i).getCell(0).getStringCellValue();
		}catch (Exception e) {
		}
		if(temp.equals(t)){
			if(custome_name_temp.equals(cpy_name)){
				time++;				
				indexs.add(i);
			}
			//System.out.println(custome_name_temp);			
			//System.out.println(temp);
			//System.out.println(t);
		}
	}
	if(time>0){
		addSubHeader(doc, t);
		String service = "";
		if(sheet.getSheetName().equals("机柜服务")){
			service = "机柜服务客户";
		}else if(sheet.getSheetName().equals("机房服务")){
			service = "机房服务客户";
		}else if(sheet.getSheetName().equals("CTP软件服务")){
			service = "CTP软件服务客户";
		}else if(sheet.getSheetName().equals("CTPmini")){
			service = "CTP MINI服务客户";
		}
		String subtitle = " 贵公司作为"+service+"，我司于"+sDate+"至"+eDate+"期间共对贵司提供"+time+"次有流程记录的"+t+"服务，具体服务内容如下：";
		addBoldContent(doc, subtitle);
		addEnter(doc);
		for(Iterator<Integer> it=indexs.iterator();it.hasNext();){
			int i=it.next();
			String content = sheet.getRow(i).getCell(3).getStringCellValue();
			addContent(doc, content);
			addEnter(doc);
		}	
		return true;
	}else{
		//addBoldContent(doc, "贵公司本年度未有服务记录，感谢贵公司对我司该项服务的一贯支持和信任。");
		return false;
	}
	//System.out.println(time);	
}

//标题1
private void addHeader(XWPFDocument doc, String text){
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	pa.setStyle("1");
	run = pa.createRun();
	run.setText(text);
//	run.setFontFamily("宋体");
//	run.setFontSize(14);
//	run.setBold(true);
	
}

//标题2
private void addSubHeader(XWPFDocument doc, String text){
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	pa.setStyle("2");
	run = pa.createRun();
	run.setText(text);
//	run.setFontFamily("宋体");
//	run.setFontSize(12);
//	run.setBold(true);
}

//内容
private void addContent(XWPFDocument doc, String text){
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	pa.setStyle(styles[1]);
	run = pa.createRun();
	run.setText(text);
//	run.setFontFamily("宋体");
//	run.setFontSize(12);
	run.setBold(false);
}

//加粗题内容
private void addBoldContent(XWPFDocument doc, String text){
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	pa.setStyle(styles[2]);
	run = pa.createRun();
	run.setText(text);
//	run.setFontFamily("宋体");
//	run.setFontSize(12);
	run.setBold(true);
}

private void addEnter(XWPFDocument doc){
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	run = pa.createRun();
	run.addCarriageReturn();
}

private void addEnding(XWPFDocument doc) {
	XWPFParagraph pa;
	XWPFRun run;
	pa = doc.createParagraph();
	pa.setStyle(styles[0]);
	run = pa.createRun();	
	run.setText("感谢贵公司对我司一贯的信赖与支持!");
//	run.setFontFamily("Arial");
//	run.setFontSize(16);
//	run.setBold(false);
}

private int findIndex(String name_temp) {
		
		int i=0;
		for(;i<10;i++){
			if(sheet_names[i].equals(name_temp)){
				break;				
			}
		}
		return i;
	}

	
}
