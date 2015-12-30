package com.sfit.doc;

import java.util.ArrayList;
import java.util.List;

import com.sfit.sheet.ServiceContent;

/**
 * 服务报告内容
 * @author chen.guoshu
 * 20150818
 */
public class ReportContent {

	/**
	 * 服务报告类型
	 * 2.年报
	 * 1.季报
	 */
	private int type = 1;
	//服务内容
	private List<ServiceContent> serviceContentLst;
	//期货公司名称
	private String company;
	
	public ReportContent(){
		serviceContentLst = new ArrayList<ServiceContent>();
	}
	
	public int getType() {
		return type;
	}
	public void setType(int type) {
		this.type = type;
	}
	public List<ServiceContent> getServiceContentLst() {
		return serviceContentLst;
	}
	public void setServiceContentLst(List<ServiceContent> serviceContentLst) {
		this.serviceContentLst = serviceContentLst;
	}
	public String getCompany() {
		return company;
	}
	public void setCompany(String company) {
		this.company = company;
	}	
}
