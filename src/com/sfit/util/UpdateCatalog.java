package com.sfit.util;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 更新word目录
 * @author chen.guoshu
 *
 */
public class UpdateCatalog {

	private UpdateCatalog(){
	}
	
	static{
		System.loadLibrary("jacob-1.17-M2-x64");
		 
	}
	
	public static void update(String filepath){
		File dir = new File(filepath);
		File[] files = dir.listFiles();
		if(files == null)
			return;
		for(int i=0; i<files.length; i++){
			if(!files[i].isDirectory()){
				updateOne(filepath+ "/"  +files[i].getName());
			}
		}
	}
	
	public static void updateOne(String fileName){

		// 初始化com的线程，非常重要！！使用结束后要调用 realease方法
		ComThread.InitSTA();
         // Instantiate objWord //Declare word object
        ActiveXComponent objWord =  new  ActiveXComponent( "Word.Application" );
         // Assign a local word object
        Dispatch wordObject = (Dispatch) objWord.getObject();
         // Create a Dispatch Parameter to show the document that is opened
        Dispatch.put((Dispatch) wordObject,  "Visible" ,  new  Variant( false )); // new
                                                                             // Variant(true)表示word应用程序可见
         // Instantiate the Documents Property
        Dispatch documents = objWord.getProperty( "Documents" ).toDispatch();  // documents表示word的所有文档窗口，（word是多文档应用程序）
		

        Dispatch document = Dispatch.call(documents,  "Open" ,  fileName ).toDispatch();;
		
		/**获取目录*/
		Dispatch tablesOfContents = Dispatch.get(document,"TablesOfContents").toDispatch();
		
			
			/**获取第一个目录。若有多个目录，则传递对应的参数*/
			Variant tablesOfContent = Dispatch.call(tablesOfContents, "Item", new Variant(1)); 
			 
			/**更新目录，有两个方法：Update　更新域，UpdatePageNumbers　只更新页码*/
			Dispatch toc = tablesOfContent.toDispatch();
			Dispatch.call(toc, "Update");
		

		/**保存word文档*/
		Dispatch.call(document, "Save");
		
        // end for
       Dispatch.call(objWord,  "Quit" );

		// 释放com线程。根据jacob的帮助文档，com的线程回收不由java的垃圾回收器处理
		ComThread.Release(); 
	}
	
/*	public static void main(String[] args) {
		//updateOne("F:/2014年第二季度对外服务报告(安粮期货).docx");
		updateOne("F:/test/2014年第二季度对外服务报告(安粮期货).docx");
	}*/
}
