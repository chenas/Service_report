package com.sfit.frame;

import java.awt.LayoutManager;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JPanel;
import javax.swing.JTextField;

import com.sfit.util.Profile;

public abstract class ReportPane extends JPanel {

	private static final long serialVersionUID = 1L;

	protected JButton startDateLb = null;
	protected JTextField startDateTxtf = null;

	protected JButton endDateLb = null;
	protected JTextField endDateTxtf = null;

	protected JButton yearBtn = null;
	protected JTextField yearTxtf = null;
	
	protected JButton openSrcBtn = null;
	protected JTextField srcPathTxtf = null;

	protected JButton openTemplateBtn = null;
	protected JTextField templatePathTxtf = null;

	protected JButton generateBtn = null;
	protected JButton updateCatalogBtn = null;

	protected String fileInFolder = "";

	public ReportPane(LayoutManager layout) {
		super(layout);
		initUI();
	}

	protected void initUI() {
		startDateLb = new JButton("开始日期");
		add(startDateLb);
		startDateTxtf = new JTextField("2015年1月1日");
		add(startDateTxtf);

		endDateLb = new JButton("结束日期");
		add(endDateLb);
		endDateTxtf = new JTextField("2015年12月31日");
		add(endDateTxtf);

		yearBtn = new JButton("年份");
		add(yearBtn);
		yearTxtf = new JTextField("2015");
		add(yearTxtf);

		setUI();

		openSrcBtn = new JButton("选择数据源");
		openSrcBtn.addActionListener(new OpenSrcBtnListener());
		add(openSrcBtn);
		srcPathTxtf = new JTextField();
		add(srcPathTxtf);
		openTemplateBtn = new JButton("打开模板");
		openTemplateBtn.addActionListener(new OpenTemplateBtnListener());
		add(openTemplateBtn);
		templatePathTxtf = new JTextField();
		add(templatePathTxtf);
		
		generateBtn = new JButton("开始生成报表");
		generateBtn.addActionListener(new GenerateBtnListener());
		add(generateBtn);
		
		updateCatalogBtn = new JButton("开始更新目录");
		updateCatalogBtn.addActionListener(new UpdateCatalogBtnListener());
		add(updateCatalogBtn);
	}

	// 用于子类继承后，增加控件
	protected void setUI() {
	}

	// 打开文件事件
	protected void openEvent(int docType) {
		JFileChooser jfc = new JFileChooser(Profile.lastPath);
		int openresult = jfc.showOpenDialog(null);
		if (openresult == JFileChooser.APPROVE_OPTION) {
			switch (docType) {
			case 1:
				srcPathTxtf.setText(jfc.getSelectedFile().getAbsolutePath());
				break;
			case 2:
				templatePathTxtf.setText(jfc.getSelectedFile().getAbsolutePath());
				break;
			default:
				break;
			}
			System.out.println(jfc.getSelectedFile().getPath());
			fileInFolder = jfc.getCurrentDirectory().toString();
			Profile.write(jfc.getSelectedFile().getParent());
		}
	}

	class OpenSrcBtnListener implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			openEvent(1);
		}
	}

	class OpenTemplateBtnListener implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			openEvent(2);
		}
	}
	
	//生成报表
	protected abstract void generateEvent();
	//更新目录
	protected abstract void updateCatalogEvent();
	
	class GenerateBtnListener implements ActionListener{
		public void actionPerformed(ActionEvent e) {
			generateEvent();
		}		
	}
	
	class UpdateCatalogBtnListener implements ActionListener{
		public void actionPerformed(ActionEvent e) {
			updateCatalogEvent();
		}
	}
}
