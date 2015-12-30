package com.sfit.sheet;

public class Summary {

	private String company;
	
	private String TITLE_11 = "总体运行情况描述";
	
	private String TITLE_CONTENT_11 = "尊敬的客户，您在yyyy年度所采购的我司的服务项分别为ssss，这些服务yyyy年度的总体运行情况均达到了相应合同条款的承诺，感谢您在yyyy年度对于我司工作的配合，希望您在nnnn年度继续与我司精诚合作，为相关服务的稳定可靠运行共同努力！";
	
	private String TITLE_12 = "满意度调查结果分析";
	
	private String summary;

	public String getCompany() {
		return company;
	}

	public void setCompany(String company) {
		this.company = company;
	}

	public String getSummary() {
		return summary;
	}

	public void setSummary(String summary) {
		this.summary = summary;
	}

	public String getTITLE_11() {
		return TITLE_11;
	}

	public void setTITLE_11(String tITLE_11) {
		TITLE_11 = tITLE_11;
	}

	public String getTITLE_CONTENT_11() {
		return TITLE_CONTENT_11;
	}

	public void setTITLE_CONTENT_11(String tITLECONTENT_11) {
		TITLE_CONTENT_11 = tITLECONTENT_11;
	}

	public String getTITLE_12() {
		return TITLE_12;
	}

	public void setTITLE_12(String tITLE_12) {
		TITLE_12 = tITLE_12;
	}
		
}
