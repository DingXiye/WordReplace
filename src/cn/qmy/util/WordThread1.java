package cn.qmy.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.Format;
import java.util.Iterator;
import java.util.List;

import javax.swing.JTextArea;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.spire.doc.Document;
import com.spire.doc.DocumentObject;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.BreakType;

public class WordThread1 implements Runnable {
	private XWPFDocument xwpf;
	private String filePath;
	private String destPath;
	private String oldStrr;
	private int page;
	private JTextArea ta;
	long startTime;
	long endTime;
	private boolean Changeflag;// 判断是否有替换词

	public WordThread1(String filePath, String destPath, int page, String oldStrr, JTextArea ta) {
		this.filePath = filePath;
		this.destPath = destPath;
		this.oldStrr = oldStrr;
		this.page = page;
		this.ta = ta;
		this.Changeflag = false;
	}

	public boolean copy(int count) {
		ta.append("复制" + count + "页任务开始..." + "\r\n");
		startTime = System.currentTimeMillis();
		// 加载文档1
		Document doc1 = new Document();
		File srcFile = new File(filePath);
		if (!srcFile.exists()) {
			ta.append(filePath + "未找到，请检查路径!!!");
			return false;
		}
		doc1.loadFromFile(filePath);
		
		// 加载文档2
		Document target = new Document();
		
		destPath="C:\\Users\\dingye\\Desktop\\2.docx";
		File targetFile = new File(destPath);
		if (!targetFile.exists()) {
			try {
				targetFile.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
				ta.append(destPath + "创建失败!!!\r\n");
				return false;
			}
			target.loadFromFile(destPath);
			Section sec= target.getSections().get(0);
	        //插入分页符
	        sec.getParagraphs().get(0).appendBreak(BreakType.Page_Break);
			// 遍历文档1中的所有子对象
			for (int i = 0; i < doc1.getSections().getCount(); i++) {
				System.out.println(i);
				Section section = doc1.getSections().get(i);
				for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++) {
					Object object = section.getBody().getChildObjects().get(j);
					// 复制文档1中的正文内容添加到文档2
					target.getSections().get(0).getBody().getChildObjects().add(((DocumentObject) object).deepClone());
				}
			}
			// 保存文档2
			target.saveToFile(destPath, FileFormat.Docx_2013);
			target.dispose();
		}

		ta.append("第1页复制完成..." + "\r\n");
		Document document = null;
		for (int i = 1; i < count; i++) {
			// 加载第一个文档
			document = new Document(destPath);
			// 使用insertTextFromFile方法将第二个文档的内容插入到第一个文档
			document.insertTextFromFile(filePath, FileFormat.Docx_2013);
			ta.append("第" + (i + 1) + "页复制完成..." + "\r\n");
			document.saveToFile(destPath, FileFormat.Docx_2013);
		}
		ta.append("复制" + count + "页任务完成\r\n");
		return true;
	}

	/**
	 * 替换所有表格中的某个字段，同时按照顺序进行编号
	 * 
	 * @param filePath
	 *            文档路径
	 * @param orderNum
	 *            设置需要读取的第几个表格(暂时未用到)
	 * @param oldStr
	 *            要替换的字段
	 * @param replaceType
	 *            替换字段的格式
	 */
	public void tableInWord(Integer orderNum, String oldStr, String replaceType) {
		ta.append("开始替换任务，替换" + oldStr + "\r\n");
		FileOutputStream outStream = null;
		FileInputStream in = null;
		String[] oldStrs = oldStr.split("\\s+");
		try {
			in = new FileInputStream(destPath);// 载入文档
			// 处理docx格式 即office2007以后版本
			if (destPath.toLowerCase().endsWith("docx")) {
				// word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
				xwpf = new XWPFDocument(in);// 得到word文档的信息
				Iterator<XWPFTable> it = xwpf.getTablesIterator();// 得到word中的表格
				int total = 0;// 记录表格的数量
				while (it.hasNext()) {
					XWPFTable table = it.next();
					System.out.println("这是第" + total + "个表的数据");
					List<XWPFTableRow> rows = table.getRows();
					// 读取每一行数据
					for (int i = 0; i < rows.size(); i++) {
						XWPFTableRow row = rows.get(i);
						// 读取每一列数据
						List<XWPFTableCell> cells = row.getTableCells();
						for (int j = 0; j < cells.size(); j++) {
							XWPFTableCell cell = cells.get(j);
							System.out.print(cell.getText() + "[" + i + "," + j + "]" + "\t");
							if (cell.getText().contains(("：" + oldStrs[0]))
									|| cell.getText().contains(("：" + oldStrs[1]))) {
								Changeflag = true;
								total++;
								List<XWPFParagraph> paragraphs = cell.getParagraphs();
								for (XWPFParagraph xwpfParagraph : paragraphs) {
									List<XWPFRun> xwpfRuns = xwpfParagraph.getRuns();
									for (int index = 0; index < xwpfRuns.size(); index++) {
										XWPFRun run = xwpfRuns.get(index);
										String runStr = run.getText(0);
										for (String old : oldStrs) {
											if (runStr.contains(old)) {
												String prefixstr = runStr.split(old)[0];
												// 实例化format，格式为“000”
												Format f1 = new DecimalFormat(replaceType);
												// 将1变为001
												String count = f1.format(total);
												runStr.replace(runStr, count);
												runStr = prefixstr + count;
												run.setText(runStr, 0);
											}
										}
									}
								}
							}
						}
						System.out.println();
					}
				}
				outStream = new FileOutputStream(destPath);
				xwpf.write(outStream);
				outStream.close();
			}
			if (Changeflag) {
				ta.append("替换任务完成.\r\n");
			} else {
				ta.append("未找到替换词" + oldStr + ".\r\n");
			}
			endTime = System.currentTimeMillis();
			long useTime = (endTime - startTime) / 1000;
			ta.append("本次任务共花费" + useTime + "s.\r\n");
			ta.append("=============================\r\n");
		} catch (Exception e) {
			e.printStackTrace();
			ta.append("替换任务失败，请重新替换。\r\n");
		} finally {
			if (xwpf != null) {
				try {
					xwpf.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	@Override
	public void run() {
		boolean flag = copy(page);
		if (flag) {
			tableInWord(2, oldStrr, "000");
		}
	}

}
