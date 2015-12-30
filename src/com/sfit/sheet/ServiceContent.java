package com.sfit.sheet;

import java.util.List;
import java.util.Map;

/**
 * 对应excel中服务各项内容
 * @author chen.guoshu
 * 20150818
 */
public class ServiceContent {

	/**服务目录 */
	private String catalog;
	/**关注点分类*/
	private Map<List<String>, List<String>> categoryDescMap;
	/**归属的期货公司*/
	private String company;
	
	public String getCatalog() {
		return catalog;
	}
	public void setCatalog(String catalog) {
		this.catalog = catalog;
	}
	public String getCompany() {
		return company;
	}
	public void setCompany(String company) {
		this.company = company;
	}
	public Map<List<String>, List<String>> getCategoryDescMap() {
		return categoryDescMap;
	}
	public void setCategoryDescMap(Map<List<String>, List<String>> categoryDescMap) {
		this.categoryDescMap = categoryDescMap;
	}
	
}
