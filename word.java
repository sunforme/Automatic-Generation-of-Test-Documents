package main;

import java.awt.Dimension;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.RandomAccessFile;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.Document;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import existdb.XqueryFile;

public class word2 {
	/**
	 * 根据模板生成新word文档 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
	 * 
	 * @param inputUrl
	 *            模板存放地址
	 * @param outputUrl
	 *            新文档存放地址
	 * @param textMap
	 *            需要替换的信息集合
	 * @param tableList
	 *            需要插入的表格信息集合
	 * @return 成功返回true,失败返回false
	 * @throws IOException
	 */
	static int table7count = 4;

	public static void templet(String fileName, String savepath) throws IOException {
		// 模板文件地址
		String inputUrl = "X:\\template\\report template_V2.0.docx";
		// 新生产的模板文件
		String outputUrl = MainUI.savePath + "/" + fileName + "_测试报告.docx";
		String outputUrl_1 = "X:\\template\\t1_report template_V1.0.docx";
		String template_1 = "X:\\template\\Detailed Question_report template_V1.0.docx";
		String template_2 = "X:\\template\\NOTEPart_report template_V1.0.docx";// Note报告模板
		String outputUrl_2 = "X:\\template\\t4_report template_V1.0.docx";

		XqueryFile xq = new XqueryFile();
		String user = "admin";
		String pwd = "123";

		String query1;
		String result1;


		String query_count = "";
		String result_count = "";

		Map<String, String> testMap = new HashMap<String, String>();

		// 从txt取xml到word的映射关系,第一部分：开头~表4，版本V1.0,Data Mapping XML to
		// Report_part1_V1.0.txt
		File f1 = new File("X:\\template\\Data Mapping XML to Report_part1_V2.0.txt");
		BufferedReader bf1 = new BufferedReader(new FileReader(f1));
		String str1;
		while ((str1 = bf1.readLine()) != null) {
			String[] s1 = str1.split(",");
			// System.out.println("path----------------" + s1[0]);

			// 结果
			query1 = "for $result in \r\n" + "doc('/db/产品测试数据集/" + fileName + "/" + "xmlfile" + ".xml')" + s1[0]
					+ "\r\n return \r\n" + " data($result)";
			 System.out.println(query1);
			xq.QueryXML(query1, user, pwd);
			result1 = xq.re;
			// 表
			testMap.put(s1[1], result1);
		}
		System.out.println("word第一部分信息填入完成");
		
		// 表格数据计算部分
		// 从txt取xml到word的映射关系,第四部分：需要计算的基础数据，版本V1.0,Data Mapping XML to
		// Report_part4_V1.0.txt
		File f4 = new File("X:\\template\\Data Mapping XML to Report_part4_V1.0.txt");
		BufferedReader bf4 = new BufferedReader(new FileReader(f4));
		String str4;
		while ((str4 = bf4.readLine()) != null) {
			String[] s = str4.split("，");

			// 结果
			query_count = "count(doc('/db/产品测试数据集/" + fileName + "/" + "xmlfile" + ".xml')" + s[0] + ")";
			xq.QueryXML(query_count, user, pwd);
			result_count = xq.re;
			// 表
			testMap.put(s[1], result_count);
		}

		NumberFormat format = NumberFormat.getPercentInstance();
		format.setMaximumFractionDigits(2);// 设置保留几位小数
		// 表41计算部分
		String result4102 = testMap.get("4102");
		String result4101 = testMap.get("4101");
		String result4105 = testMap.get("4105");
		String result4104 = testMap.get("4104");
		float result4103 = Float.parseFloat(result4102) / Float.parseFloat(result4101);
		float result4106 = Float.parseFloat(result4105) / Float.parseFloat(result4104);
		int result4107 = Integer.parseInt(result4101) + Integer.parseInt(result4104);
		int result4108 = Integer.parseInt(result4102) + Integer.parseInt(result4105);
		float result4109 = (float) result4108 / (float) result4107;
		// 表41结果填入
		testMap.put("4103", format.format(result4103));
		testMap.put("4106", format.format(result4106));
		testMap.put("4107", Integer.toString(result4107));
		testMap.put("4108", Integer.toString(result4108));
		testMap.put("4109", format.format(result4109));

		

		// txt路径
		String txtpath_tlist = "X:\\template\\problem information_V2.0.txt";
		FileInputStream f_tlist = new FileInputStream(txtpath_tlist);
		InputStreamReader isr_tlist = new InputStreamReader(f_tlist, "UTF-8");
		BufferedReader bf_tlist = new BufferedReader(isr_tlist);
		String str_tlist;

		List<String[]> testList = new ArrayList<String[]>();
		while ((str_tlist = bf_tlist.readLine()) != null) {
			String[] s_tlist = str_tlist.split("\\|");

			testList.add(new String[] { s_tlist[0], s_tlist[1], s_tlist[2], "/" });
		}

		String file = "";
		for (int i = 1; i < 21; i++) {
			String q = "for $result in \r\n" + "doc('/db/产品测试数据集/" + fileName + "/" + "xmlfile" + ".xml')"
					+ "/测试/被测件基本信息/测试依据/依据[" + i + "]\r\n return \r\n" + " data($result)";
			xq.QueryXML(q, user, pwd);
			String r = xq.re;
			if (r.length() > 0) {
				file = file + r + "、";
			}
		}
		file = file.substring(0, file.length() - 1);
		System.out.println("file：" + file);
		testMap.put("pfile1", file);

		word2.changWord(inputUrl, outputUrl_1, testMap, testList, fileName);

		try {
			ProblemTable_Create.copytable(outputUrl_1, outputUrl_2, template_1);
			ProblemTable_Create.copyNote(outputUrl_2, outputUrl, template_2);
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("word生成成功");

	}

	public static boolean changWord(String inputUrl, String outputUrl, Map<String, String> textMap,
			List<String[]> tableList, String fileName) {

		// 模板转换默认成功
		boolean changeFlag = true;
		try {
			// 获取docx解析对象
			XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
			// 解析替换文本段落对象
			word2.changeText(document, textMap, fileName);
			// 解析替换表格对象
			word2.changeTable(document, textMap, tableList, fileName);

			// 生成新的word
			File file = new File(outputUrl);
			FileOutputStream stream = new FileOutputStream(file);
			document.write(stream);
			stream.close();

		} catch (IOException e) {
			e.printStackTrace();
			changeFlag = false;
		}

		return changeFlag;

	}

	/**
	 * 替换段落文本
	 * 
	 * @param document
	 *            docx解析对象
	 * @param textMap
	 *            需要替换的信息集合
	 */
	public static void changeText(XWPFDocument document, Map<String, String> textMap, String fileName) {
		// 获取段落集合
		List<XWPFParagraph> paragraphs = document.getParagraphs();

		for (XWPFParagraph paragraph : paragraphs) {
			// 判断此段落时候需要进行替换
			String text = paragraph.getText();
			if (checkText(text)) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (XWPFRun run : runs) {
					// 替换模板原来位置
					run.setText(changeValue(run.toString(), textMap), 0);
				}
			}
		}
	}

	/**
	 * 替换表格对象方法
	 * 
	 * @param document
	 *            docx解析对象
	 * @param textMap
	 *            需要替换的信息集合
	 * @param tableList
	 *            需要插入的表格信息集合
	 */
	public static void changeTable(XWPFDocument document, Map<String, String> textMap, List<String[]> tableList,
			String fileName) {
		// 获取表格对象集合
		List<XWPFTable> tables = document.getTables();
		for (int i = 0; i < tables.size(); i++) {
			// 只处理行数大于等于2的表格，且不循环表头
			XWPFTable table = tables.get(i);
			if (table.getRows().size() > 1) {
				// 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
				if (checkText(table.getText())) {
					List<XWPFTableRow> rows = table.getRows();
					// 遍历表格,并替换模板
					eachTable(rows, textMap, fileName);
				} else {
					insertTable(table, tableList);
				}
			}
		}
	}

	/**
	 * 遍历表格
	 * 
	 * @param rows
	 *            表格行对象
	 * @param textMap
	 *            需要替换的信息集合
	 */
	public static void eachTable(List<XWPFTableRow> rows, Map<String, String> textMap, String fileName) {

		for (XWPFTableRow row : rows) {
			List<XWPFTableCell> cells = row.getTableCells();
			for (XWPFTableCell cell : cells) {
				// 判断单元格是否需要替换
				if (checkText(cell.getText())) {
					XqueryFile xq = new XqueryFile();
					String user = "admin";
					String pwd = "123";

					} // file结束


					List<XWPFParagraph> paragraphs = cell.getParagraphs();
					for (XWPFParagraph paragraph : paragraphs) {
						List<XWPFRun> runs = paragraph.getRuns();
						for (XWPFRun run : runs) {
							run.setText(changeValue(run.toString(), textMap), 0);

						}
					}
				

			}
		}
	}

	/**
	 * 为表格插入数据，行数不够添加新行
	 * 
	 * @param table
	 *            需要插入数据的表格
	 * @param tableList
	 *            插入数据集合
	 */
	public static void insertTable(XWPFTable table, List<String[]> tableList) {
		// 创建行,根据需要插入的数据添加新行，不处理表头
		for (int i = 1; i < tableList.size(); i++) {
			XWPFTableRow row = table.createRow();
		}
		// 遍历表格插入数据
		List<XWPFTableRow> rows = table.getRows();
		for (int i = 1; i < rows.size(); i++) {
			XWPFTableRow newRow = table.getRow(i);
			List<XWPFTableCell> cells = newRow.getTableCells();
			for (int j = 0; j < cells.size(); j++) {
				XWPFTableCell cell = cells.get(j);

				List<XWPFParagraph> paragraphs = cell.getParagraphs();
				// System.out.println("cell的paragraph的数量是："+ paragraphs.size());
				for (XWPFParagraph paragraph : paragraphs) {

					XWPFRun headRun = paragraph.createRun();
					// headRun.setBold(bold);// 是否粗体
					headRun.setText(tableList.get(i - 1)[j], 0);
					headRun.setFontSize(12);

				}

				if (j != 1) {
					// 设置水平居中,需要ooxml-schemas包支持
					CTTc cttc = cell.getCTTc();
					CTTcPr ctPr = cttc.addNewTcPr();
					ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
					cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
				}

			}
		}

	}

	/**
	 * 判断文本中时候包含$
	 * 
	 * @param text
	 *            文本
	 * @return 包含返回true,不包含返回false
	 */
	public static boolean checkText(String text) {
		boolean check = false;
		if (text.indexOf("$") != -1 && text != "${file}") {
			check = true;
		}
		return check;

	}

	/**
	 * 匹配传入信息集合与模板
	 * 
	 * @param value
	 *            模板需要替换的区域
	 * @param textMap
	 *            传入信息集合
	 * @return 模板需要替换区域信息集合对应值
	 */
	public static String changeValue(String value, Map<String, String> textMap) {
		Set<Entry<String, String>> textSets = textMap.entrySet();
		for (Entry<String, String> textSet : textSets) {
			// 匹配模板与替换值 格式${key}
			String key = "${" + textSet.getKey() + "}";
			if (value.indexOf(key) != -1) {
				value = textSet.getValue();
			}
		}
		// 模板未匹配到区域替换为空
		if (checkText(value)) {
			value = "";
		}
		return value;
	}

}
