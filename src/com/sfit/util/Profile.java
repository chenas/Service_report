package com.sfit.util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.Properties;

/**
 * 保存上次打开的路径
 * 
 * @author chen.guoshu
 * 
 */
public class Profile {
	public static String lastPath = "C:/";
	static String file = "/config.properties";
	static Properties pros = null;
	
	static {
		if (pros == null)
			pros = new Properties();
		InputStream inp = Profile.class.getResourceAsStream(file);
		try {
			pros.load(inp);
			lastPath = pros.getProperty("lastPath").trim();
		} catch (IOException e) {
			System.out.println(e.toString());
		}
	}

	private Profile() {
	}

	// 保存每次打开文件的路径
	public static boolean write(String latestPath) {
		lastPath = latestPath;
		return writeCfg("lastPath", latestPath);		
	}

	// 读取配置项
	public static String getPara(String key) {
		return pros.getProperty(key);
	}

	// 保存配置项
	public static boolean writeCfg(String key, String value) {
		OutputStream os = null;
		boolean b = true;
		System.out.println(Profile.class.getResource(file).getPath());
		try {
			os = new FileOutputStream(Profile.class.getResource(file).getPath());
			pros.setProperty(key, value);
			pros.store(os, new Date().toString()); // 将属性写入
			os.flush();
			os.close();
		} catch (IOException ioe) {
			b = false;
			ioe.printStackTrace();
		}
		return b;
	}
	
	public static void main(String[] args){
		System.out.println(writeCfg("test", "中文"));
		System.out.println(getPara("SheetName"));
	}
}
