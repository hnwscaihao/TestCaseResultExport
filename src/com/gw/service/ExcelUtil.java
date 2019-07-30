package com.gw.service;

import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Map.Entry;
import java.util.Set;

import javax.imageio.ImageIO;
import javax.print.attribute.standard.OutputDeviceAssigned;
import javax.swing.JOptionPane;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.lf5.viewer.categoryexplorer.CategoryExplorerTree;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.gw.ui.TestResultExportUI;
import com.gw.util.GenerateXmlUtil;
import com.gw.util.MKSCommand;
import com.mks.api.response.APIException;

@SuppressWarnings("all")
public class ExcelUtil {
	private static final String POST_CONFIG_FILE = "FieldMapping.xml";
	private static final String CATEGORY_CONFIG_FILE = "Category.xml";
	Map<String, List<Map<String, String>>> xmlConfig = new HashMap<>();
	Map<String, List<String>> contentColumns = new HashMap<>();
	private List<String> contentHeaders = new ArrayList<>();
	private List<String> stepHeaders = new ArrayList<>();
	private List<String> realStepFields = new ArrayList<>();
	private List<String> resultHeaders = new ArrayList<>();
	private Map<String, String> inputCountMap = new HashMap<String, String>();
	private String xmlType = "Test Suite";
	private String SEQUENCE = "Sequence";
	private List<List<Object>> datas = new ArrayList<>();
	private List<List<String>> listHeaders = new ArrayList<>();
	private List<CellRangeAddress> cellList = new ArrayList<>();
	public static final Map<String, String> HEADER_MAP = new HashMap<String, String>();
	public static final Map<String, String> HEADER_COLOR_RECORD = new HashMap<String, String>();
	private Map<String, String> stepHeaderMap = new HashMap<String, String>();
	private Map<String, String> resultHeaderMap = new HashMap<String, String>();
	private List<String> headers = new ArrayList<>();// 第一行标题
	private List<String> headerTwos = new ArrayList<>();// 第二行标题
	private List<Object> data = new ArrayList<>();
	private static Integer startInteger = 0;
	private static Integer resultStartIndex = 0;
	private static Integer InputOutputRange = 5;
	private static Integer EInputOutputRange = 5;
	private static Integer CycleTest = 1;
	private static Integer Depth = 1;
	private List<String> testResultHeaders = new ArrayList<>();
	private Map<String, Object> testResultData = new HashMap<>();
	private boolean hasStepField;
	private boolean hasResultField;
	public static final List<String> CURRENT_CATEGORIES = new ArrayList<String>();// 记录导入对象的正确Category

	public static final Map<String, List<String>> PICK_FIELD_RECORD = new HashMap<String, List<String>>();

	public static final Map<String, String> FIELD_TYPE_RECORD = new HashMap<String, String>();

	public static final Map<String, String> IMPORT_DOCUMENT_TYPE = new HashMap<String, String>();

	/**
	 * 解析XML
	 * 
	 * @param project
	 * @throws APIException
	 */
	public List<String> parseXML() throws APIException {
		List<String> exportTypes = new ArrayList<String>();// 导出类型list
		try {
			TestResultExportUI.logger.info("start to parse xml : " + POST_CONFIG_FILE);

			Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder()
					.parse(ExcelUtil.class.getClassLoader().getResourceAsStream(POST_CONFIG_FILE));
			Element root = document.getDocumentElement();

			if (root != null) {
				NodeList eleList = root.getElementsByTagName("importType");// 获取导出类型
				if (eleList != null) {
					List<Map<String, String>> allFields = new ArrayList<>(); // 模板field
					List<String> ptcFields = new ArrayList<>();// 存放系统自带field
					for (int i = 0; i < eleList.getLength(); i++) {// 获取当前类型节点相关属性值
						Element item = (Element) eleList.item(i);
						String typeName = item.getAttribute("name");
						exportTypes.add(typeName);
						String documentType = item.getAttribute("type");
						IMPORT_DOCUMENT_TYPE.put(typeName, documentType);
						parseData(item, allFields, ptcFields);// 解析数据，往Excel模板汇中存放field
					}
					xmlConfig.put(xmlType, allFields);
					contentColumns.put(xmlType, ptcFields);
				}
			}
		} catch (ParserConfigurationException e) {
			TestResultExportUI.logger.error("parse config file exception", e);
		} catch (SAXException e) {
			TestResultExportUI.logger.error("get config file exception", e);
		} catch (IOException e) {
			TestResultExportUI.logger.error("io exception", e);
		} finally {
			TestResultExportUI.logger.info(" xmlConfig: " + xmlConfig + " \n, the ptcTestCaseColumns: " + contentColumns);
			return exportTypes;
		}
	}

	/**
	 * Description 查询当前要导入类型的 正确Category
	 * 
	 * @param documentType
	 * @throws Exception
	 */
	public void parseCurrentCategories(String documentType) throws Exception {
		Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
				.parse(ExcelUtil.class.getClassLoader().getResourceAsStream(CATEGORY_CONFIG_FILE));
		Element root = doc.getDocumentElement();
		List<String> typeList = new ArrayList<String>();
		// 得到xml配置
		NodeList importTypes = root.getElementsByTagName("documentType");
		for (int j = 0; j < importTypes.getLength(); j++) {
			Element importType = (Element) importTypes.item(j);
			String typeName = importType.getAttribute("name");
			if (typeName.equals(documentType)) {
				NodeList categoryNodes = importType.getElementsByTagName("category");
				for (int i = 0; i < categoryNodes.getLength(); i++) {
					Element categoryNode = (Element) categoryNodes.item(i);
					CURRENT_CATEGORIES.add(categoryNode.getAttribute("name"));
				}
			}
		}
	}

	/**
	 * 导出TestSuite对象到 Excel模板
	 * 
	 * @param tObjIds
	 * @param cmd
	 * @param path
	 * @throws Exception
	 */
	public void exportReport(List<String> tObjIds, MKSCommand cmd, String path) throws Exception {
		this.parseXML(); // 解析xml
		int resultStep = resultHeaders.size();
		GenerateXmlUtil.caseHeaders.addAll(new ArrayList<String>(contentHeaders));
		// 这里查询testSession 字段拼接
		// caseHeaders.add("Test Environment");
		int parentIndex = -1;
		int maxLength = 0;
		for (String id : tObjIds) {
			List<String> caseFields = contentColumns.get(xmlType);
			caseFields.add("Contains");
			caseFields.add("Test Step");//合并单元格用
			List<Map<String, String>> testCaseItem = cmd.allContents(id, caseFields);// 测试用例字段
			List<String> testStepFields = cmd.getTestStepFields();// testStep字段集合
			List<Map<String, String>> list = this.xmlConfig.get(xmlType);
			int col = 0;
			for (String field : contentHeaders) {
				headerTwos.add("");
				headers.add(field);
				CellRangeAddress input = new CellRangeAddress(0, 1, col, col);// 单元格范围
				col++;
				cellList.add(input);
			}

			int dataIndex = 0;
			String sub_sequence = null;
			String sub_sub_sequence = null;
			for (Map<String, String> testCase : testCaseItem) {
				// Test Case 中的字段
				data = new ArrayList<>();// 行数据
				for (String header : contentHeaders) {
					String realField = HEADER_MAP.get(header);
					String value = realField == null ? ""
							: testCase.get(realField) == null ? "" : testCase.get(realField);
					if (SEQUENCE.equals(header)) {//截取Sequence
						if(value.indexOf("_")>-1){
							sub_sequence = value.substring(value.indexOf("_"));
							value = value.substring(0, value.indexOf("_"));
							if(sub_sequence.indexOf("_")>-1){
								sub_sub_sequence = sub_sub_sequence.substring(sub_sub_sequence.indexOf("_"));
								sub_sequence = sub_sequence.substring(0, sub_sequence.indexOf("_"));
							}
						}
					}
					data.add(value);
				}

				// ---------Test Step
				if (hasStepField) {
					List<String> StepsIDsList = new ArrayList<>();// testStepiD集合
					String steps = testCase.get("Test Steps");
					if (steps != null && !"".equals(steps)) { // 如果Test
																// Steps字段不为空
																// 再去查里面的字段。
						String[] StepsID = steps.split(",");
						for (int i = 0; i < StepsID.length; i++) {
							StepsIDsList.add(StepsID[i]);
						}
						getStepsItem(cmd, StepsIDsList, testStepFields, headers);// 再次查询Test
																						// Steps
																						// 中字段。
					} // steps
				}

				// 获得TestResult信息。
				if (resultStartIndex < data.size()) {
					resultStartIndex = data.size();
				}
				if (maxLength < data.size()) {
					maxLength = data.size();
				}
				datas.add(data);// 拼接完所有数据 为一行
			} // testCaseFor
			/** 空行时，添加testStep 列标题 */
			if (hasStepField && !(headerTwos.contains("Test Step ID") || headerTwos.contains("Input I/F")
					|| headerTwos.contains("Output I/F"))) {
				if (stepHeaders.contains("Source Filename")) {
					headers.add("Call Depth");
					headers.add("-");
					headers.add("-");
					headers.add("-");
					headers.add("-");
					headers.add("-");
				} else {
					headers.add("Test Step"); // 如果有添加进 第一行标题
					headers.add("-");
					headers.add("-");
					headers.add("-");
				}
				for (String stepField : stepHeaders) {
					headerTwos.add(stepField);
				}
				if (stepHeaders.contains("Source Filename")) {
					CellRangeAddress Output = new CellRangeAddress(0, 0, contentHeaders.size(),
							contentHeaders.size() + 5);
					cellList.add(Output);
				} else {// MCU 合并4个单元格
					CellRangeAddress Output = new CellRangeAddress(0, 0, contentHeaders.size(),
							contentHeaders.size() + 3);
					cellList.add(Output);
				}
				maxLength = contentHeaders.size() + stepHeaders.size();
				if (startInteger < maxLength) {
					startInteger = maxLength;
				}
			}

			if (hasResultField) {// 含有测试结果列
				for (int k = 0; k < testCaseItem.size(); k++) {
					Map<String, String> testCase = testCaseItem.get(k);
					List<Object> data = datas.get(k);
					int emptyCount = maxLength - data.size();
					for (int m = 0; m < emptyCount; m++) {
						data.add("");
					}
					String structureVal = (String) data.get(parentIndex);
																				// 导出，P
																				// 结构且有子级不导出测试结果
					String containValue = testCase.get("Contains");
					getTestResult(cmd, testCase, data);
				}
			}

			listHeaders.add(headers);// 添加完第一行标题
			listHeaders.add(headerTwos);// 添加完第二行标题

		}

		if (hasResultField && !(headerTwos.contains("Severity") || headerTwos.contains("Reproducibility")
				|| headerTwos.contains("Tester"))) {
			headers.add("1-Cycle Test");
			for (String stepField : resultHeaders) {
				headerTwos.add(stepField);
				headers.add(" ");
			}
			headers.remove(headers.size() - 1);
			if (startInteger < contentHeaders.size()) {
				startInteger = contentHeaders.size();
			}
			CellRangeAddress result = new CellRangeAddress(0, 0, startInteger, startInteger + resultHeaders.size() - 1);
			cellList.add(result);
			maxLength = maxLength + resultHeaders.size();
		}
		for (List<Object> rowData : datas) {
			int emptyCellCount = headers.size() - rowData.size();
			if (emptyCellCount > 0) {
				for (int count = 0; count < emptyCellCount; count++) {
					rowData.add("");
				}
			}
		}
		List<String> needMoreWidthField = new ArrayList<String>();
		needMoreWidthField.add("Text(Description)");
		needMoreWidthField.add("Expected Results");
		needMoreWidthField.add("Test Environment");
		needMoreWidthField.add("Test case description");
		needMoreWidthField.add("Test Case");
		needMoreWidthField.add("Input I/F & Data");
		needMoreWidthField.add("Output I/F & Data");

		/** 获取Category信息 */
		String documentType = "Test Suite";
		parseCurrentCategories(documentType);
		/** 获取Category信息 */
		/** 获取 Pick 值信息 */
		List<String> importFields = new ArrayList<String>();
		for (String header : contentHeaders) {
			String field = HEADER_MAP.get(header);
			if (!"-".equals(field)) {
				importFields.add(field);
			}
		}
		FIELD_TYPE_RECORD.putAll(cmd.getAllFieldType(importFields, PICK_FIELD_RECORD));
		replaceLogid(cmd);// logid 替换为FullName(工号)

		String trueType = "Test Case & Result";
		/** 获取 Pick 值信息 */
		Workbook wookbook = GenerateXmlUtil.exportComplexExcel(listHeaders, datas, needMoreWidthField, trueType,
				cellList);
		if (wookbook == null) {
			wookbook = new HSSFWorkbook();
		}
		// 拼接判断文件路径名称
		String documentName = MyRunnable.class.newInstance().documentName;
		SimpleDateFormat df = new SimpleDateFormat("yyyy_MM_dd");// 设置日期格式
		String time = df.format(new Date());
		try {
			String actualPath = path.endsWith(".xls") ? path
					: path + "\\" + documentName + "_" + time + "_(" + 1 + ")-" + tObjIds.get(0).toString()
							+ ".xls";
			File file = new File(actualPath);
			if (!file.exists()) {
				outputFromwrok(actualPath, wookbook);
			} else {
				int showConfirmDialog = JOptionPane.showConfirmDialog(TestResultExportUI.contentPane,
						"The file already exists, Whether to overwrite this file?");
				String absolutePath = file.getAbsolutePath();
				if (showConfirmDialog == 0) {// 覆盖
					outputFromwrok(actualPath, wookbook);
				} else if (showConfirmDialog == 1) {// 不覆盖
					File pathFile = new File(actualPath);
					String parent = pathFile.getParent();
					File fileDir = new File(parent);
					if (fileDir.isDirectory()) {
						File[] listFiles = fileDir.listFiles();
						int count = 1;
						for (File file2 : listFiles) {
							String filePath = file2.toString();
							if (filePath.endsWith("xls") || filePath.endsWith("xlsx")) {
								if ((filePath.endsWith("-" + tObjIds.get(0).toString() + ".xls")
										|| filePath.endsWith("-" + tObjIds.get(0).toString() + ".xlsx"))
										&& filePath.contains(documentName + "_" + time)) {
									count++;
								}
							}
						}
						String actualPath2 = path + "\\" + documentName + "_" + time + "_(" + count + ")-"
								+ tObjIds.get(0).toString() + ".xls";
						outputFromwrok(actualPath2, wookbook);
					}

				}

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * @param cmd
	 * @throws APIException
	 */
	private void replaceLogid(MKSCommand cmd) throws APIException {
		// logid 替换为FullName(工号)
//		for (List<Object> list : datas) {
//			for (int p = 0; p < list.size(); p++) {
//				Object obj = list.get(p);
//				if ( obj != null && ( obj.toString().matches("[G][W]\\d{0,9}")
//						||  obj.toString().matches("[g][w]\\d{0,9}"))) {
//					String fullName = cmd.getUserNames(obj.toString());
//					if (null != fullName && fullName.length() > 0) {
//						list.set(p, fullName + "（" + list.get(p) + "）");
//					}
//
//				}
//			}
//		}
	}

	/**
	 * 解析数据，Excel模板中存放field
	 * 
	 * @param eleList
	 * @param list
	 * @param ptcFields
	 */
	private void parseData(Element exportType, List<Map<String, String>> list, List<String> ptcFields) {
		xmlType = exportType.getAttribute("name");
		NodeList nodeFields = exportType.getElementsByTagName("excelField");
		for (int j = 0; j < nodeFields.getLength(); j++) {
			Map<String, String> map = new HashMap<>();// 存放所有fields Excel模板
			Element fields = (Element) nodeFields.item(j);
			String field = fields.getAttribute("field");
			String type = fields.getAttribute("type");
			String fieldName = fields.getAttribute("name");
			String titleColor = fields.getAttribute("titleColor");// 获取标题颜色标识
			if (!field.equals("-") && !type.equals("Test Result") && !type.equals("Test Step")) {
				ptcFields.add(field);// 如果模板中符合以上情况，则直接将field存放到系统自带field的list中
			}
			if ("Test Step".equals(type)) {
				hasStepField = true;
				stepHeaders.add(fieldName);
				String parentField = fields.getAttribute("parentField");
				if (parentField != null && !realStepFields.contains(parentField))
					realStepFields.add(parentField);
				stepHeaderMap.put(fieldName, field);
				HEADER_COLOR_RECORD.put(fieldName, titleColor);
			} else if ("Test Result".equals(type)) {
				hasResultField = true;
				if (!resultHeaders.contains(fieldName))
					resultHeaders.add(fieldName);
				resultHeaderMap.put(fieldName, field);
			} else {
				if (!contentHeaders.contains(fieldName)) {
					contentHeaders.add(fieldName);
					HEADER_MAP.put(fieldName, field);
					HEADER_COLOR_RECORD.put(fieldName, titleColor);
				}
			}
			map.put("name", fields.getAttribute("name"));
			list.add(map);
		}
	}

	/**
	 * 测试结果处理方法
	 * 
	 * @param cmd
	 * @param testCase
	 * @throws APIException
	 */
	private boolean getTestResult(MKSCommand cmd, Map<String, String> testCase, List<Object> data)
			throws APIException {
		List<Map<String, Object>> result = cmd.getResult(testCase.get("ID"), testCase.get("ID"), "Test Case");
		if (result != null && result.size() > 0) {
			if (CycleTest < result.size()) {
				CycleTest = result.size();
			}
			int i = 1;
			int step = resultHeaders.size() - 1;
			if (startInteger < contentHeaders.size()) {
				startInteger = contentHeaders.size();
			}
			String mergeName = "Cycle Test";
			int testEnvironIndex = headers.indexOf("Test Environment");
			for (Map<String, Object> map : result) {
				Object sessionID = map.get("sessionID");
				int testerIndex = 0;
				int testDateIndex = 0;
				for (String field : resultHeaders) {
					String realField = resultHeaderMap.get(field);
					data.add(map.get(realField));
					if ("Tester".equals(field)) {
						testerIndex = data.size() - 1;
					}
					if ("Date of Test".equals(field) || "Test Date".equals(field)) {
						testDateIndex = data.size() - 1;
					}
				}

				// 查询Test Session
				if (inputCountMap.get(i + mergeName) == null) {
					headers.add(i + mergeName);
					for (int m = 0; m < step; m++) {
						headers.add("-");
					}
					for (String field : resultHeaders) {
						headerTwos.add(field);// 加二级标题
					}
					CellRangeAddress input = new CellRangeAddress(0, 0, startInteger, startInteger + step);
					cellList.add(input);
					startInteger = startInteger + step + 1;
					inputCountMap.put((i) + mergeName, "Result");
				}
				i++;
			}
		} // result
		return false;
	}

	/**
	 * 获取Test Steps中的字段
	 * 
	 * @param cmd
	 * @param StepsIDsList
	 * @param testSteps
	 * @param headers
	 * @param data
	 * @throws APIException
	 */
	private void getStepsItem(MKSCommand cmd, List<String> StepsIDsList, List<String> testSteps, List<String> headers
			) throws APIException {
		List<Map<String, String>> itemByIds = cmd.getItemByIds(StepsIDsList, testSteps);
		List<String> stepsData = new ArrayList<>();
		int i = 1;
		int step = stepHeaders.size();
		if (startInteger == 0)
			startInteger = contentHeaders.size();
		for (Map<String, String> map : itemByIds) {
			if (map.get("ID") != null) {// 对test Steps 做截取
				String ID = map.get("ID");
				if (realStepFields.contains("Call Depth")) {
					String CallDepth = map.get("Call Depth");
					if(CallDepth!=null && !CallDepth.equals("")){
						data.add(ID);
					}
					
				} else {
					for (String field : stepHeaders) {
						String realField = stepHeaderMap.get(field);
						String val = realField == null ? "" : map.get(realField) == null ? "" : map.get(realField);
						data.add(val);
					}
					if (inputCountMap.get(i + "-Test Step") == null) {
						headers.add("Test Step"); // 如果有添加进 第一行标题
						for (int j = 0; j < step - 1; j++) {
							headers.add("-");
						}
						for (String field : stepHeaders) {
							headerTwos.add(field);
						}
						// 处理合并单元格
						CellRangeAddress Output = new CellRangeAddress(0, 0, startInteger, startInteger + step - 1);
						cellList.add(Output);
						// 处理headers中的位置
						startInteger = startInteger + step;
						inputCountMap.put(i + "-Test Step", "add");
					}
				}
			}
			i++;// 第几次循环
		}
	}

	public static void outputFromwrok(String filePath, Workbook wookbook) {
		try {
			FileOutputStream output = new FileOutputStream(filePath);
			wookbook.write(output);
			output.flush();
			TestResultExportUI.class.newInstance().isParseSuccess = true;

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 设置Border
	 * 
	 * @param style
	 * @param top
	 * @param bottom
	 * @param left
	 * @param right
	 * @param border
	 */
	public static void setBorder(HSSFCellStyle style, boolean top, boolean bottom, boolean left, boolean right,
			short border) {
		if (top)
			style.setBorderTop(border);
		if (bottom)
			style.setBorderBottom(border);
		if (left)
			style.setBorderLeft(border);
		if (right)
			style.setBorderRight(border);
	}

}
