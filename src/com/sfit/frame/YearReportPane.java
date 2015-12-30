package com.sfit.frame;

import java.awt.LayoutManager;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import com.sfit.doc.GenerateRepoter;
import com.sfit.util.UpdateCatalog;

public class YearReportPane extends ReportPane{
	private static final long serialVersionUID = 1L;

	public YearReportPane(LayoutManager layout) {
		super(layout);
	}
	
	@Override
	protected void setUI() {
	}

	@Override
	protected void generateEvent() {
		// 开始生成文件
		String filepath = this.srcPathTxtf.getText();
		String templatePath = this.templatePathTxtf.getText();
		String startDate = this.startDateTxtf.getText();
		String endDate = this.endDateTxtf.getText();
		generateBtn.setText("正在生成报表...");
		GenerateRepoter gr = new GenerateRepoter();
		gr.GenerateReport(Integer.parseInt(this.yearTxtf.getText()),filepath,fileInFolder,templatePath, startDate, endDate, "");
		generateBtn.setText("重新生成报表");
		JFrame msg = new JFrame();
		JOptionPane.showMessageDialog(msg, "生成完成！");
		msg.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	@Override
	protected void updateCatalogEvent() {
		updateCatalogBtn.setText("正在更新报表目录...");
		String folder = fileInFolder+"/生成文件";
		UpdateCatalog.update(folder);
		updateCatalogBtn.setText("重新更新报表目录");
		JFrame msg = new JFrame();
		JOptionPane.showMessageDialog(msg, "更新完成！");
		msg.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}



}
