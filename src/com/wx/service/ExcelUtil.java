package com.wx.service;

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

import com.mks.api.response.APIException;
import com.wx.ui.TestResultExportUI;
import com.wx.util.GenerateXmlUtil;
import com.wx.util.MKSCommand;

@SuppressWarnings("all")
public class ExcelUtil {
	private static final String POST_CONFIG_FILE = "FieldMapping.xml";
	private static final String CATEGORY_CONFIG_FILE = "Category.xml";
	Map<String, List<Map<String, String>>> xmlConfig = new HashMap<>();
	Map<String, List<String>> contentColumns = new HashMap<>();
	private List<String> allHeaders = new ArrayList<>();
	private List<String> contentHeaders = new ArrayList<>();
	private List<String> stepHeaders = new ArrayList<>();
	private List<String> realStepFields = new ArrayList<>();
	private List<String> resultHeaders = new ArrayList<>();
	private Map<String, String> inputCountMap = new HashMap<String, String>();
	private String xmlType = "Test Suite";
	private String SEQUENCE = "Sequence";
	private List<List<Object>> datas = new ArrayList<>();
	private List<List<List<Object>>> allDatas = new ArrayList<>();
	private List<List<String>> listHeaders = new ArrayList<>();
	private List<CellRangeAddress> cellList = new ArrayList<>();
	public static final Map<String, String> HEADER_MAP = new HashMap<String, String>();
	public static final Map<String, String> HEADER_COLOR_RECORD = new HashMap<String, String>();
	private Map<String, String> stepHeaderMap = new HashMap<String, String>();
	private Map<String, String> resultHeaderMap = new HashMap<String, String>();
	private List<String> headers = new ArrayList<>();// 第一行标题
	private List<String> headerTwos = new ArrayList<>();// 第二行标题
	private Map<String,List<String>> allSheetHeaders = new HashMap<>();
	private List<Object> data = new ArrayList<>();
	private static Integer CycleTest = 1;
	private List<String> testResultHeaders = new ArrayList<>();
	private Map<String, Object> testResultData = new HashMap<>();
	private String suitId;
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
				stepHeaders.add(fieldName);
				realStepFields.add(fieldName);
				stepHeaderMap.put(fieldName, field);
				HEADER_COLOR_RECORD.put(fieldName, titleColor);
				if(!headers.contains("Test Steps")){
					headers.add("Test Steps");
				}else
					headers.add("-");
				headerTwos.add(fieldName);
			} else if ("Test Result".equals(type)) {
				if (!resultHeaders.contains(fieldName))
					resultHeaders.add(fieldName);
				resultHeaderMap.put(fieldName, field);
				if(!headers.contains("1-Cycle Test")){
					headers.add("1-Cycle Test");
				}else
					headers.add("-");
				headerTwos.add(fieldName);
			} else {
				if (!contentHeaders.contains(fieldName)) {
					contentHeaders.add(fieldName);
					headers.add(fieldName);
					headerTwos.add("-");
					HEADER_MAP.put(fieldName, field);
					HEADER_COLOR_RECORD.put(fieldName, titleColor);
				}
			}
			map.put("name", fields.getAttribute("name"));
			list.add(map);
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
		int dataIndex = 0;
		List<String> firstHeaders = null;
		List<String> secondHeaders = null;
		List<String> needMoreWidthField = new ArrayList<String>();
		needMoreWidthField.add("Summary");
		needMoreWidthField.add("Expected Results");
		needMoreWidthField.add("Test");
		needMoreWidthField.add("Description");

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
//		replaceLogid(cmd);// logid 替换为FullName(工号)

		String sheetName = "Test Cases";
		/** 获取 Pick 值信息 */
		HSSFWorkbook wookbook = new HSSFWorkbook();
		for (String suitId : tObjIds) {
			List<String> caseFields = contentColumns.get(xmlType);
			caseFields.add("Contains");
			caseFields.add("Test Steps");//合并单元格用
//			List<Map<String, String>> testCaseItem = cmd.allContents(suitId, caseFields);// 测试用例字段
			List<List<Map<String,String>>> allTestCaseItems = cmd.allContentsByHeading(suitId, caseFields);//测试用例字段
			List<String> testStepFields = cmd.getTestStepFields();// testStep字段集合
			List<Map<String, String>> list = this.xmlConfig.get(xmlType);
			int col = 0;
			for(List<Map<String,String>> testCaseItems : allTestCaseItems){//根据一级Heading拆分写入不同Sheet
				cellList = new ArrayList<>();
				for(int i=0; i<headers.size(); i++){
					String headerTwo = headerTwos.get(i);
					CellRangeAddress input = null;
					if("-".equals(headerTwo)){//上下合并单元格
						input = new CellRangeAddress(0, 1, i, i);
					}else if(i < headers.size() - 1){
						int temp = 1;
						String header = headers.get(i + temp);
						while("-".equals(header)){
							temp++;
							if( i + temp == headers.size()){
								break;
							}
							header = headers.get(i + temp);
						}
						input = new CellRangeAddress(0, 0, i, i + temp - 1);//首行合并单元格
					}
					if(input != null)
						cellList.add(input);
				}
				if(!testCaseItems.isEmpty()){
					Map<String,String> firstCase = testCaseItems.get(0);
					String sequence = firstCase.get(HEADER_MAP.get(SEQUENCE));
					sheetName = sequence == null || "".equals(sequence)? sheetName : sequence;
				}
				datas = new ArrayList<>();
				firstHeaders = new ArrayList<>(headers);
				secondHeaders = new ArrayList<>(headerTwos);
				listHeaders = new ArrayList<>();
				for (Map<String, String> testCase : testCaseItems) {
					// Test Case 中的字段
					String sub_sequence = null;
					String sub_sub_sequence = null;
					data = new ArrayList<>(headers.size());// 行数据
					datas.add(data);// 拼接完所有数据 为一行
					for (String header : headers) {
						if(contentHeaders.contains(header)){
							String realField = HEADER_MAP.get(header);
							String value = realField == null ? ""
									: testCase.get(realField) == null ? "" : testCase.get(realField);
							if (SEQUENCE.equals(header)) {//截取Sequence
								if(value.indexOf("_")>-1){
									sub_sequence = value.substring(value.indexOf("_")+1);
									value = value.substring(0, value.indexOf("_"));
									if(sub_sequence.indexOf("_")>-1){
										sub_sub_sequence = sub_sequence.substring(sub_sequence.indexOf("_")+1);
										sub_sequence = sub_sequence.substring(0, sub_sequence.indexOf("_"));
									}
								}
							}else if("Sub-Sequence".equals(header))
								value = sub_sequence;
							else if("Sub-Sub-Sequence".equals(header))
								value = sub_sub_sequence;
							data.add(headers.indexOf(header), value);
						}else
							data.add(headers.indexOf(header), "");
						
						// 根据Test Step 合并单元格
					}
					dataIndex = datas.size() + 2;//当前用例列的数据行数
					// ---------Test Step
					List<String> StepsIDsList = new ArrayList<>();// testStepiD集合
					String steps = testCase.get("Test Steps");
					if (steps != null && !"".equals(steps)) { // 如果Test Steps字段不为空  再去查里面的字段。
						String[] StepsID = steps.split(",");
						
						for (int i = 0; i < StepsID.length; i++) {
							StepsIDsList.add(StepsID[i]);
							/** 补充 N个 test Step的行出来用来合并单元格*/
							List<Object> stepRowData = new ArrayList<>(headers.size());// 行数据
							if(i > 0){//第二个Test Step开始添加行
								for(int m=0; m<headers.size(); m++){
									stepRowData.add(m,"");
								}
								datas.add(stepRowData);
							}
							/** 补充 N个 test Step的行出来用来合并单元格*/
						}
						for(int m=0; m<headerTwos.size(); m++){// 添加用例数据合并单元格
							if(!stepHeaders.contains(headerTwos.get(m))){
								CellRangeAddress input = new CellRangeAddress(dataIndex - 1 , dataIndex + StepsID.length - 2, m, m );
								cellList.add(input);
							}
						}
						getStepsItem(cmd, StepsIDsList, testStepFields, dataIndex - 2);// 再次查询Test Steps 中字段。
					}
					
					//导出 测试结果数据
					getTestResult(cmd, testCase, data, firstHeaders, secondHeaders);
				} 
				listHeaders.add(firstHeaders);// 添加完第一行标题
				listHeaders.add(secondHeaders);// 添加完第二行标题
				GenerateXmlUtil.exportComplexExcel(wookbook, listHeaders, datas, needMoreWidthField, sheetName,
						cellList);
			}
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
		for (List<Object> list : datas) {
			for (int p = 0; p < list.size(); p++) {
				Object obj = list.get(p);
				String fullName = cmd.getUserNames(obj.toString());
				if (null != fullName && fullName.length() > 0) {
					list.set(p, fullName + "（" + list.get(p) + "）");
				}

			}
		}
	}

	/**
	 * 测试结果处理方法
	 * 
	 * @param cmd
	 * @param testCase
	 * @throws APIException
	 */
	private boolean getTestResult(MKSCommand cmd, Map<String, String> testCase, List<Object> data, List<String> firstHeaders, List<String> secondHeaders)
			throws APIException {
		List<Map<String, Object>> result = cmd.getResult(testCase.get("ID"), testCase.get("ID"), "Test Case");
		if (result != null && result.size() > 0) {
			if (CycleTest < result.size()) {
				CycleTest = result.size();
			}
			int i = 1;
			String mergeName = "Cycle Test";
			int testEnvironIndex = headers.indexOf("Test Environment");
			for (Map<String, Object> map : result) {
				Object sessionID = map.get("sessionID");
				Object verdict = map.get("verdict");
				Object annotation = map.get("Annotation");
				if(i == 1){
					data.set(headerTwos.indexOf("Session ID"), sessionID);
					data.set(headerTwos.indexOf("Verdict"), verdict);
					data.set(headerTwos.indexOf("Annotation"), annotation);
				}else{
					data.add(sessionID);
					data.add(verdict);
					data.add(annotation);
				}

				if (inputCountMap.get(i + mergeName) == null) {
					if(i > 2){
						firstHeaders.add(i + mergeName);
						firstHeaders.add("-");
						firstHeaders.add("-");
						for (String field : resultHeaders) {
							secondHeaders.add(field);// 加二级标题
						}
						CellRangeAddress input = new CellRangeAddress(0, 0, firstHeaders.size() - 3 , firstHeaders.size() - 1);
						cellList.add(input);
						inputCountMap.put((i) + mergeName, "Result");
					}
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
	private void getStepsItem(MKSCommand cmd, List<String> StepsIDsList, List<String> testSteps, Integer startIndex
			) throws APIException {
		List<Map<String, String>> itemMaps = cmd.getItemByIds(StepsIDsList, testSteps);
		List<String> stepsData = new ArrayList<>();
		for(int count = 0; count < itemMaps.size(); count++){
			Map<String, String> map  = itemMaps.get(count);
			List<Object> rowData = datas.get(startIndex + count - 1);
			for (String header : stepHeaders) {
				String realField = stepHeaderMap.get(header);
				String val = realField == null ? "" : map.get(realField) == null ? "" : map.get(realField);
				if(count == 0){
					rowData.set(headerTwos.indexOf(header),val);//封装数据
				}else{
					rowData.set(headerTwos.indexOf(header),val);//封装数据
				}
			}
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
