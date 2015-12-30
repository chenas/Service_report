package com.sfit.doc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * 写服务报告
 * @author chen.guoshu
 * 20150818
 */
public class WriterUtil {

	//字体style
	private final static String TITLE_1_STYLE = "1";
	private final static String TITLE_2_STYLE = "2";
	private final static String CONTENT_STYLE = "a4";
	private final static String END_STYLE = "ac";

	/**
	 * 
	 * 写入WORD
	 * @param doc
	 * @param text
	 * @param style
	 * @param isBold
	 */
	public static void addSubject(XWPFDocument doc, String text, String style, boolean isBold){
		XWPFParagraph pa;
		XWPFRun run;
		pa = doc.createParagraph();
		pa.setStyle(style);
		run = pa.createRun();
		run.setText(text);
		run.setBold(isBold);
	}

	/**
	 * 
	 * 写入一级标题
	 * 对应excel中的服务目录
	 * @param doc
	 * @param text
	 */
	public static void addCatalog(XWPFDocument doc, String text){
		addSubject(doc, text, TITLE_1_STYLE, false);
	}
	
	/**
	 * 
	 * 写入二级标题
	 * 对应excel中的分类
	 * @param doc
	 * @param text
	 */
	public static void addCategory(XWPFDocument doc, String text){
		addSubject(doc, text, TITLE_2_STYLE, false);
	}
	
	/**
	 * 写入内容
	 * @param doc
	 * @param text
	 */
	public static void addContent(XWPFDocument doc, String text){
		addSubject(doc, text, CONTENT_STYLE, false);
	}
	
	/**
	 * 写入粗体内容
	 * @param doc
	 * @param text
	 */
	public static void addBoldContent(XWPFDocument doc, String text){
		addSubject(doc, text, CONTENT_STYLE, true);
	}

	/**
	 * 写入结束语
	 * @param doc
	 * @param text
	 */
	public static void addEnd(XWPFDocument doc, String text){
		addSubject(doc, text, END_STYLE, false);
	}
	
	/**
	 * 回车
	 * @param doc
	 */
	public static void addEnter(XWPFDocument doc){
		XWPFParagraph pa;
		XWPFRun run;
		pa = doc.createParagraph();
		run = pa.createRun();
		run.addCarriageReturn();
	}
	
	public static void main(String[] args){
		InputStream inp = null;
		XWPFDocument doc = null;
		OutputStream os = null;
		try {
			inp = new FileInputStream("D:/SFIT/服务台2014年报/对外服务报告模版（内部版）.docx");
			os = new FileOutputStream("D:/SFIT/new.docx");
			doc = new XWPFDocument(inp);
			

			String[] styles = new String[6];
			List<XWPFParagraph> paras = doc.getParagraphs();
			System.out.println(paras.size());
			for(int j=1;j<7;j++){
				XWPFParagraph xp1;
				xp1 = paras.get(paras.size()-j) ;
				//System.out.println(xp1.getParagraphText());
				System.out.println(xp1.getStyle());
				styles[j-1] = xp1.getStyle();				
			}			
			
			addCatalog(doc, "标题1");
			addCategory(doc, "标题2");
			addContent(doc, "内容列出段");
			addBoldContent(doc, "加粗的内容");
			addEnter(doc);
			addBoldContent(doc, "加粗内容3饿2222");
			addContent(doc, "内容的的点点的");
			addEnd(doc, "内容的的点点的");
			doc.write(os);
			os.flush();
			os.close();
		} catch (FileNotFoundException e) {
			System.out.println(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println(e.toString());
			e.printStackTrace();
		}
	}
}
