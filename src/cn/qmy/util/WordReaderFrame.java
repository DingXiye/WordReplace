package cn.qmy.util;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import cn.qmy.mail.SendMail;
import cn.qmy.mail.UtilMethod;

/**
 *
 * 读取word文档中表格数据，支持仅支持docx doc文件没有版本进行测试
 * 先扩展指定页数，然后替换扩展文档中的所有标签
 * @author dingye
 *
 */
public class WordReaderFrame {
	private String filePath;// 文档路径
	private String destPath;
	JTextArea ta = new JTextArea();

	public static void main(String[] args) {
		WordReaderFrame wordUtil = new WordReaderFrame();
		wordUtil.guiFrom();
	}

	/**
	 * 界面
	 */
	public void guiFrom() {
		final JFrame frame = new JFrame("甜儿专属");
		frame.setSize(410, 520);
		frame.setLocation(700, 200);
		frame.setLayout(null);

		JPanel pInput = new JPanel();
		pInput.setBounds(10, 10, 375, 170);
		pInput.setLayout(new GridLayout(5, 2, 10, 10));

		JButton startButton = new JButton("开始替换");
		startButton.setBounds(10, 120 + 10, 100, 20);

		JButton openFileButton = new JButton("选择本地文件");
		openFileButton.setBounds(10, 120 + 30, 100, 20);
		// openFileButton.setEnabled(false);
		final JTextField pathText = new JTextField();
		pathText.setEditable(false);

		JLabel changeStrL = new JLabel("替换词(两个词请用空格隔开):");
		final JTextField changeStrT = new JTextField();

		JLabel toPageL = new JLabel("扩展页数:");
		final JTextField toPageT = new JTextField();

		JLabel targetPathL = new JLabel("目标路径:");
		final JTextField targetPathT = new JTextField();
		targetPathT.setEditable(false);

		pInput.add(openFileButton);
		pInput.add(pathText);

		pInput.add(targetPathL);
		pInput.add(targetPathT);

		pInput.add(changeStrL);
		pInput.add(changeStrT);

		pInput.add(toPageL);
		pInput.add(toPageT);

		pInput.add(startButton);

		frame.add(pInput);

		// 文本域
		JLabel result = new JLabel("执行日志:");
		result.setBounds(10, 110 + 50, 80, 80);

		ta = new JTextArea();
		ta.setLineWrap(true);
		ta.setWrapStyleWord(true);// 激活断行不断字功能

		JScrollPane scroll = new JScrollPane(ta);

		// 分别设置水平和垂直滚动条自动出现
		scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		scroll.setBounds(10, 180 + 30, 370, 230);
		scroll.setVisible(true);

		frame.add(result);
		frame.add(scroll);

		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		openFileButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				final JFileChooser jfc = new JFileChooser();// 文件选择器
				jfc.showOpenDialog(null);
				File file = jfc.getSelectedFile();
				if (file != null) {
					filePath = file.getAbsolutePath();
					destPath = filePath.split("\\.docx")[0] + "bobo.docx";
					pathText.setText(filePath);
					targetPathT.setText(destPath);
				} else {
					filePath = null;
					pathText.setText("未选择文件");
				}
			}
		});
		pathText.setText(filePath);
		targetPathT.setText(destPath);
		// 先将源文件扩展为多页，再替换扩展文件中的词
		startButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// 1、生成随机随机验证码
				List<String> results = new UtilMethod().genCodes(6, 1);
				System.out.println(results);
				// 2、将验证码发送到163邮箱
				boolean flag = new SendMail().sendMails("dyxiye@163.com", results.get(0));
				// 3、验证验证码
//				if (flag) {
//					Object obj = JOptionPane.showInputDialog(null, "请输入验证码：\n", "验证码", JOptionPane.PLAIN_MESSAGE, null,
//							null, "在这输入");
//					if (obj.equals(results.get(0))) {
						int page = Integer.parseInt(toPageT.getText());
						String oldStr = changeStrT.getText();
						WordThread wordThread = new WordThread(filePath, destPath, page, oldStr, ta);
						new Thread(wordThread).start();
						ta.append("正在启动系统..." + "\r\n");
//					}else{
//						JOptionPane.showMessageDialog(null,"验证码错误", "错误信息",JOptionPane.ERROR_MESSAGE);
//						return ;
//					}
//				}else{
//					JOptionPane.showMessageDialog(null,"邮件发送失败", "错误信息",JOptionPane.ERROR_MESSAGE);
//				}
			}
		});
	}

}