package com.sfit.frame;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;

import javax.swing.JCheckBoxMenuItem;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;

public class MainFrame extends JFrame {

	private static final long serialVersionUID = 1L;
	private JTabbedPane pane = new JTabbedPane();
	private JMenuItem scrollTabItem;
	private JMenuItem componentTabItem;
	private final int numTab = 3;

	public MainFrame(String title) {
		// 设置frame标题名
		super(title);
		// 设置关闭方式
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		// 创建菜单栏
		initMenu();
		// 将tabpane添加到frame中
		add(pane);

	}

	// 初始化tabpanel相关配置，并且用于resize的调用
	public void runTest() {
		pane.removeAll();
		// 设置有关闭按键的标签
		componentTabItem.setSelected(true);
		// 创建标签
		pane.add("月 报", createContent("月 报", 0));
		pane.add("年 报", createContent("年 报", 1));
		pane.add("说 明", createContent("说 明", 2));
		// 初始化标签上的文字和Button
		initTabComponent(0);
		initTabComponent(1);
		initTabComponent(2);
		// 设置frame的大小
		setSize(new Dimension(500, 500));
		// 将frame放到屏幕的正中央
		setLocationRelativeTo(null);
		// 显示frame
		setVisible(true);

	}

	// 创建标签内容部分
	private Component createContent(String name, int type) {
		// 创建一个panel，并设置布局为一个分块
		// ReportPane panel = new ReportPane(new FlowLayout(FlowLayout.CENTER,
		// 5, 1));
		JPanel panel = null;
		switch (type) {
		case 0:
			panel = new MonthReportPane(new GridLayout(0, 2));
			break;
		case 1:
			panel = new YearReportPane(new GridLayout(0, 2));
			break;
		case 2:
			panel = new JPanel();
			JLabel lb = new JLabel("1.不支持03格式的ms office文档  \n " +
					" 2.先生成之后, 然后再点击更新目录");
			panel.add(lb);
			break;
		default:
			break;
		}
		
		return panel;
	}

	// 初始化带有关闭按钮的标签头部
	private void initTabComponent(int i) {
		// 用这个函数可以初始化标签的头部
		pane.setTabComponentAt(i, new TabComponent(pane));
	}

	// 创建菜单栏
	private void initMenu() {
		// 创建一个菜单条
		JMenuBar mb = new JMenuBar();

		// 创建重叠tab选项
		scrollTabItem = new JCheckBoxMenuItem("重叠tab");
		// 设置快捷键
		scrollTabItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S,
				InputEvent.CTRL_MASK));
		// 设置监听事件
		scrollTabItem.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				if (pane.getTabLayoutPolicy() == JTabbedPane.SCROLL_TAB_LAYOUT)
					pane.setTabLayoutPolicy(JTabbedPane.WRAP_TAB_LAYOUT);
				else
					pane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
			}
		});

		// 设置可关闭的标签的菜单
		componentTabItem = new JCheckBoxMenuItem("设置可关闭的tab");
		componentTabItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_C,
				InputEvent.CTRL_MASK));
		componentTabItem.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				for (int i = 0; i < numTab; i++) {
					if (componentTabItem.isSelected())
						initTabComponent(i);
					else
						pane.setTabComponentAt(i, null);
				}
			}
		});

		// 设置重置标签
		JMenuItem reSetItem = new JMenuItem("重置");
		reSetItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_R,
				InputEvent.CTRL_MASK));
		reSetItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				runTest();
			}
		});

		// 创建菜单
		JMenu menu = new JMenu("选项");
		// 添加菜单项
		menu.add(componentTabItem);
		//menu.add(scrollTabItem);
		menu.add(reSetItem);

		// 添加菜单
		mb.add(menu);
		// 添加菜单条（注意一个frame只能有一个菜单条，所以用set）
		setJMenuBar(mb);
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				new MainFrame("报表").runTest();
			}
		});
	}

}
