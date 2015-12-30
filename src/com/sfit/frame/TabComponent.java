package com.sfit.frame;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;

import javax.swing.AbstractButton;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.plaf.basic.BasicButtonUI;

public class TabComponent extends JPanel {
	private static final long serialVersionUID = 1L;
	// 父容器
	private final JTabbedPane pane;

	public TabComponent(final JTabbedPane pane) {
		// 设置标签头要水平的，从左向右的排列组件，并且组件的间距为0（横向和纵向）
		super(new FlowLayout(FlowLayout.LEFT, 0, 0));

		if (pane == null)
			throw new NullPointerException("pane can not be null");
		this.pane = pane;

		// 不画出panel的边界
		setOpaque(false);

		// 创建标签头的文字内容
		JLabel label = new JLabel() {
			private static final long serialVersionUID = 1L;

			// 重写方法，返回pane指定位置的名字，这里有点绕，可以好好理解理解
			// 为什么不直接设置标签的内容，而是通过重写函数来修改标签内容
			// 可能是希望业务逻辑尽量不要散落在外部
			@Override
			public String getText() {
				// 可以获得当前panel在tab中的位置
				int i = pane.indexOfTabComponent(TabComponent.this);
				if (i != -1)
					// 得到当前panel的名字（实际是tab的名字）
					return pane.getTitleAt(i);
				return null;
			}
		};

		label.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 5));
		add(label);
		// 创建关闭按钮（就是那个差按键）
		JButton button = new TabButton();
		add(button);
		setBorder(BorderFactory.createEmptyBorder(2, 0, 0, 0));
	}

	// 自定义关闭按钮
	private class TabButton extends JButton implements ActionListener {
		private static final long serialVersionUID = 1L;

		public TabButton() {
			// 设置按键的大小
			final int size = 17;
			setPreferredSize(new Dimension(size, size));
			// 设置按键的提示信息
			setToolTipText("关闭窗口");
			// 设置按键的绘制于普通按键相同
			setUI(new BasicButtonUI());
			// 不对Button进行填充，就是按键是透明的
			setContentAreaFilled(false);
			// 按键不能获得焦点
			setFocusable(false);
			// 设置按键的边框为雕刻样式
			setBorder(BorderFactory.createEtchedBorder());
			// 系统不自动绘制按键边界（这个边界在鼠标放上去之后才绘制）
			setBorderPainted(false);

			// 添加按键点击事件
			addActionListener(TabButton.this);
			// 添加鼠标事件（主要是mouseover 和 mouse exit）
			addMouseListener(mouseListener);

		}

		// 重写Button的绘制函数
		@Override
		protected void paintComponent(Graphics g) {
			super.paintComponent(g);

			// 创建一个graphics2D，因为需要在Button上画差
			Graphics2D g2 = (Graphics2D) g.create();

			// 设置画笔，宽度为2
			g2.setStroke(new BasicStroke(2));
			// 设置画笔颜色
			g2.setColor(Color.BLACK);
			// 当鼠标移动到Button上时，画笔为紫色
			if (getModel().isRollover())
				g2.setColor(Color.PINK);
			// 绘制差
			int delta = 6;
			g2.drawLine(delta, delta, getWidth() - delta - 1, getHeight()
					- delta - 1);
			g2.drawLine(getWidth() - delta - 1, delta, delta, getHeight()
					- delta - 1);
			// 释放画笔资源
			g2.dispose();
		}

		public void actionPerformed(ActionEvent arg0) {
			// 当按键被点击，关闭当前标签页
			int i = pane.indexOfTabComponent(TabComponent.this);
			if (i != -1)
				pane.remove(i);
		}

		private final MouseListener mouseListener = new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent event) {
				// 鼠标移入按键，绘制按键边界
				Component c = event.getComponent();
				if (c instanceof AbstractButton)
					((AbstractButton) c).setBorderPainted(true);
			}

			@Override
			public void mouseExited(MouseEvent event) {
				// 鼠标移出按键，不绘制按键边界
				Component c = event.getComponent();
				if (c instanceof AbstractButton)
					((AbstractButton) c).setBorderPainted(false);
			}

		};

	}

}