package com.gw.util;

import java.awt.Color;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
//import org.apache.poi.hssf.usermodel.HSSFChart.HSSFChartType;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddressList;
import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.gw.service.ExcelUtil;
import com.gw.ui.TestResultExportUI;

/**
 * POI导出excel表格
 * 
 * @author lv617
 * @version 1.0
 */
@SuppressWarnings("all")
public class GenerateXmlUtil {
	private static SimpleDateFormat DATE_FORAMT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	public static List<String> caseHeaders = new ArrayList<>();

	public List<String> getCaseHeaders() {
		return caseHeaders;
	}

	public void setCaseHeaders(List<String> caseHeaders) {
		this.caseHeaders = caseHeaders;
	}

	/**
	 * 导出最精简的excel表格
	 * 
	 * @param headers
	 * @param datas
	 * @param name
	 * @return
	 */
	public static Workbook generateCreateXsl(String headers[], List<List<Object>> datas, String name) {

		if (headers == null || headers.length < 1) {
			return null;
		}
		HSSFWorkbook work = new HSSFWorkbook();
		HSSFSheet sheet = work.createSheet(name);
		sheet.setDefaultRowHeightInPoints(15);
		sheet.setDefaultColumnWidth(15);
		HSSFCellStyle style = setStyle(work, "宋体", true, null, true, true, true);
		// 设置标题行
		HSSFRow row = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			HSSFCell cell = row.createCell(i);
			HSSFRichTextString text = new HSSFRichTextString(headers[i]);
			cell.setCellStyle(style);
			cell.setCellValue(text);
		}
		// 循环一次是一行
		for (List<Object> data : datas) {
			row = sheet.createRow(datas.indexOf(data) + 1);// 获取一行
			for (int i = 0; i < data.size(); i++) {
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(style);
				GenerateXmlUtil.setValue(cell, data.get(i));
			}
		}
		return work;
	}

	/**
	 * 导出复杂的excel表格(自定义格式)
	 * 
	 * @param listHeaders
	 *            复杂表头数据(结合合并单元格使用,实现复杂表头)
	 * @param datas
	 *            表格内容数据
	 * @param name
	 *            表格中sheet的名称
	 * @param merges
	 *            合并单元格; 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
	 * @param subheads
	 *            将表格数据设置为副表头; 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
	 * @param hFontName
	 *            表格表头字体格式
	 * @param dFontName
	 *            表格内容字体格式
	 * @param hSize
	 *            表格表头字体大小
	 * @param dSize
	 *            表格内容字体大小
	 * @param hBold
	 *            表头字体加粗
	 * @param dBold
	 *            内容字体加粗
	 * @param hCenter
	 *            表头字体居中
	 * @param dCenter
	 *            内容字体居中
	 * @param hForegroundColor
	 *            表头背景色; 例:红色,黄色,绿色,蓝色,紫色,灰色.
	 * @param dForegroundColor
	 *            内容背景色; 例:红色,黄色,绿色,蓝色,紫色,灰色.
	 * @param hWrapText
	 *            表头自动换行
	 * @param dWrapText
	 *            内容自动换行
	 * @param borderBottom
	 *            表格添加边框
	 * @param autoWidths
	 *            设置自适应列宽的最大展示宽度
	 * @param columnWidths
	 *            自定义列宽设置; 例:参数1：哪一列 参数2：列宽值
	 * @param rowHeights
	 *            自定义行高设置; 例:参数1：哪一行 参数2：行高值
	 * @return
	 */
	public static Workbook exportComplexExcel(List<List<String>> listHeaders, List<List<Object>> datas, String name,
			List<Integer[]> merges, List<Integer[]> subheads, String hFontName, String dFontName, Integer hSize,
			Integer dSize, boolean hBold, boolean dBold, boolean hCenter, boolean dCenter, String hForegroundColor,
			String dForegroundColor, boolean hWrapText, boolean dWrapText, boolean borderBottom, Integer autoWidths,
			List<Integer[]> columnWidths, List<Integer[]> rowHeights) {
		if (listHeaders == null || listHeaders.get(0) == null || listHeaders.size() < 1) {
			return null;
		}
		HSSFWorkbook work = new HSSFWorkbook();// 创建对象
		HSSFSheet sheet = work.createSheet(name);// 设置sheet名字
		// 设置列宽
		if (columnWidths != null && columnWidths.size() > 0 && columnWidths.get(0).length == 2) {
			for (Integer[] columnWidth : columnWidths) {
				sheet.setColumnWidth(columnWidth[0], columnWidth[1] * 256);
			}
		}
		// 设置自适应列宽
		if (autoWidths != null) {
			List<Integer> maxCalls = GenerateXmlUtil.getMaxCall(listHeaders, datas);
			for (int i = 0, j = maxCalls.size(); i < j; i++) {
				// 最大列宽设置
				if (maxCalls.get(i) > autoWidths) {
					sheet.setColumnWidth(i, autoWidths * 256);
				} else {
					sheet.setColumnWidth(i, maxCalls.get(i) * 256);
				}
			}
		}
		// 设置合并单元格
		if (merges != null && merges.size() > 0 && merges.get(0).length == 4) {
			// 设置合并单元格//参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
			for (Integer[] merge : merges) {
				CellRangeAddress region = new CellRangeAddress(merge[0], merge[1], merge[2], merge[3]);
				sheet.addMergedRegion(region);
			}
		}
		// 设置行高
		if (rowHeights != null && rowHeights.size() > 0 && rowHeights.get(0).length == 2) {
			// 设置表头格式
			// HSSFCellStyle style = GenerateXmlUtil.setStyle(work, hFontName,
			// hSize, hBold, hCenter, hForegroundColor, hWrapText,
			// borderBottom);
			// 设置 表格内容格式
			// HSSFCellStyle style2 = GenerateXmlUtil.setStyle(work, dFontName,
			// dSize, dBold, dCenter, dForegroundColor, dWrapText,
			// borderBottom);
			for (Integer[] rowHeight : rowHeights) {
				// 插入表格表头数据
				// GenerateXmlUtil.addHeader(listHeaders, sheet, style,
				// rowHeight);
				// 插入表格内容数据
				// GenerateXmlUtil.addData(listHeaders, datas, subheads, work,
				// sheet, style2, rowHeight, style);
			}
		} else {
			// 设置表头格式
			// HSSFCellStyle style = GenerateXmlUtil.setStyle(work, hFontName,
			// hSize, hBold, hCenter, hForegroundColor, hWrapText,
			// borderBottom);
			// 设置 表格内容格式
			// HSSFCellStyle style2 = GenerateXmlUtil.setStyle(work, dFontName,
			// dSize, dBold, dCenter, dForegroundColor, dWrapText,
			// borderBottom);
			// 插入表格表头数据
			// GenerateXmlUtil.addHeader(listHeaders, sheet, style, null);
			// 插入表格内容数据
			// GenerateXmlUtil.addData(listHeaders, datas, subheads, work,
			// sheet, style2, null, style);
		}
		return work;
	}

	/**
	 * 导出复杂的excel表格
	 * (复杂表头固定格式:表头字体,宋体12号,加粗,紫色背景色;内容字体,宋体10号;字体居中,自动换行,自适应列宽(最大30 ),表格加边框)
	 * 
	 * @param listHeaders
	 *            复杂表头数据(结合合并单元格使用,实现复杂表头)
	 * @param datas
	 *            表格内容数据
	 * @param name
	 *            表格中sheet的名称
	 * @param merges
	 *            合并单元格; 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
	 * @param subheads
	 *            将表格数据设置为副表头; 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
	 * @return
	 */
	public static Workbook exportComplexExcel(List<List<String>> listHeaders, List<List<Object>> datas,
			List<String> needMoreWidthField, String name, List<CellRangeAddress> merges) {

		TestResultExportUI.logger.info("reverse is data:" + datas);
		if (listHeaders == null || listHeaders.size() < 1 || listHeaders.get(0) == null) {
			return null;
		}

		HSSFWorkbook work = new HSSFWorkbook();
		HSSFSheet sheet = work.createSheet(name);
		sheet.autoSizeColumn(1, true);
		List<String> headers = listHeaders.get(0);
		List<String> headers2 = listHeaders.size() > 1 ? listHeaders.get(1) : new ArrayList<String>();
		for (int i = 0; i < headers.size(); i++) {
			String header = headers.get(i);
			String field = ExcelUtil.HEADER_MAP.get(header);
			String fieldType = ExcelUtil.FIELD_TYPE_RECORD.get(field);
			if ("longtext".equals(fieldType) || "Text".equals(field)) {
				sheet.setColumnWidth(i, 9000);
			} else if (field != null) {
				int headerLength = header.length();
				int colWidth = headerLength / 3 * 1000;
				colWidth = colWidth >= 3000 ? colWidth : 3000;// 最小宽度 2000
				sheet.setColumnWidth(i, colWidth);
			}
		}
		if (!needMoreWidthField.isEmpty()) {
			for (String field : needMoreWidthField) {
				int realIndex = 0;
				List<String> subList = new ArrayList<String>(headers2);
				int index = headers2.indexOf(field);
				while (index > -1) {
					realIndex = realIndex + index;
					subList = subList.subList(index + 1, subList.size());// 把当前位置截掉
					sheet.setColumnWidth(realIndex, 9000);
					realIndex = realIndex + 1;// 多截一个位置，补上
					index = subList.indexOf(field);
				}
			}
		}

		// 设置标题行格式
		HSSFFont titleFont = work.createFont();
		titleFont.setFontHeightInPoints((short) 10);// 字号
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 加粗
		//
		// HSSFCellStyle titleStyle = work.createCellStyle();
		// ExcelUtil.setBorder(titleStyle, true, true, true, true,
		// HSSFCellStyle.BORDER_THIN);
		// titleStyle.setWrapText(true);
		// titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		// titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 可更改字段标题样式
		HSSFCellStyle castTitleStyle = work.createCellStyle();
		ExcelUtil.setBorder(castTitleStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		castTitleStyle.setWrapText(true);
		castTitleStyle.setFont(titleFont);
		castTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		castTitleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		castTitleStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
		castTitleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 不可更改字段标题样式(系统字段)
		HSSFCellStyle unalterTitleStyle = work.createCellStyle();
		ExcelUtil.setBorder(unalterTitleStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		unalterTitleStyle.setWrapText(true);
		unalterTitleStyle.setFont(titleFont);
		unalterTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		unalterTitleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		unalterTitleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		unalterTitleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 联合评审人员未设置时，此字段可更新；若已设置，则字段不更新(初次设置有效)
		HSSFCellStyle updateTitleStyle = work.createCellStyle();
		ExcelUtil.setBorder(updateTitleStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		updateTitleStyle.setWrapText(true);
		updateTitleStyle.setFont(titleFont);
		updateTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		updateTitleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		updateTitleStyle.setFillForegroundColor(IndexedColors.LAVENDER.getIndex());
		updateTitleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 设置表格内容行格式
		HSSFCellStyle contentStyle = work.createCellStyle();
		ExcelUtil.setBorder(contentStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		contentStyle.setWrapText(true);
		contentStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		contentStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		contentStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		HSSFCellStyle passStyle = work.createCellStyle();
		ExcelUtil.setBorder(passStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		passStyle.setWrapText(true);
		passStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		passStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		passStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		HSSFCellStyle failStyle = work.createCellStyle();
		ExcelUtil.setBorder(failStyle, true, true, true, true, HSSFCellStyle.BORDER_THIN);
		failStyle.setWrapText(true);
		failStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		failStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		failStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 设置合并单元格
		if (merges != null && merges.size() > 0) {
			// 设置合并单元格//参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
			for (CellRangeAddress merge : merges) {
				sheet.addMergedRegion(merge);
			}
		}
		/** 添加数据格式验证 */
		addValidation(sheet, listHeaders, datas.size());
		// 插入表格表头数据
		GenerateXmlUtil.addHeader(listHeaders, sheet, castTitleStyle, unalterTitleStyle, updateTitleStyle, null);
		// 插入表格内容数据
		GenerateXmlUtil.addData(listHeaders, datas, null, work, sheet, contentStyle, null, passStyle, failStyle);

		return work;
	}

	/**
	 * 导出复杂的excel表格
	 * (单行表头固定格式:表头字体,宋体12号,加粗,紫色背景色;内容字体,宋体10号;字体居中,自动换行,自适应列宽(最大30 ),表格加边框)
	 * 
	 * @param headers
	 *            单行表头数据
	 * @param datas
	 *            表格内容数据
	 * @param name
	 *            表格中sheet的名称
	 * @return
	 */
	public static Workbook exportComplexExcel(String headers[], List<List<Object>> datas, String name) {
		if (headers == null || headers.length < 1) {
			return null;
		}
		HSSFWorkbook work = new HSSFWorkbook();
		HSSFSheet sheet = work.createSheet(name);
		List<String[]> listHeaders = new ArrayList<>();
		listHeaders.add(headers);
		// 设置自适应列宽
		// List<Integer> maxCalls = GenerateXmlUtil.getMaxCall(listHeaders,
		// datas);
		// for (int i = 0, j = maxCalls.size(); i < j; i++) {
		// // 最大列宽设置
		// if (maxCalls.get(i) > 30) {
		// sheet.setColumnWidth(i, 30 * 256);
		// } else {
		// sheet.setColumnWidth(i, maxCalls.get(i) * 256);
		// }
		// }
		// 设置标题行格式
		// HSSFCellStyle style = GenerateXmlUtil.setStyle(work, "宋体", 12, true,
		// true, "紫色", true, true);
		// 设置表格内容行格式
		// HSSFCellStyle style2 = GenerateXmlUtil.setStyle(work, "宋体", 10,
		// false, true, null, true, true);
		// 插入表格表头数据
		for (int m = 0, n = listHeaders.size(); m < n; m++) {
			HSSFRow row = sheet.createRow(m);
			String[] header = listHeaders.get(m);
			for (int i = 0; i < header.length; i++) {
				HSSFCell cell = row.createCell(i);
				HSSFRichTextString text = new HSSFRichTextString(header[i]);
				cell.setCellValue(text);
				// 单元格格式设置
				// cell.setCellStyle(style);
			}
		}
		// 插入表格内容数据
		for (List<Object> data : datas) {
			Integer rowNum = datas.indexOf(data) + listHeaders.size();
			HSSFRow row = sheet.createRow(rowNum);
			for (int i = 0, j = data.size(); i < j; i++) {
				HSSFCell cell = row.createCell(i);
				GenerateXmlUtil.setValue(cell, data.get(i));
				// 单元格格式设置
				// cell.setCellStyle(style2);
			}
		}
		return work;
	}

	/************************************* 手动分割线 ******************************************************/
	/**
	 * 插入表格的表头数据
	 * 
	 * @param listHeaders
	 *            复杂表头数据(结合合并单元格使用,实现复杂表头)
	 * @param sheet
	 *            HSSFSheet
	 * @param style
	 *            表格格式
	 * @param rowHeight
	 *            设置行高; 参数1：哪一行 参数2：行高值
	 */
	private static void addHeader(List<List<String>> listHeaders, HSSFSheet sheet, HSSFCellStyle castTitleStyle,
			HSSFCellStyle unalterTitleStyle, HSSFCellStyle updateTitleStyle, Integer[] rowHeight) {

		for (int m = 0, n = listHeaders.size(); m < n; m++) {
			HSSFRow row = sheet.createRow((short) m);
			if (rowHeight != null && rowHeight[0] == m) {
				row.setHeightInPoints(rowHeight[1]);
			} else {
				// 设置默认行高
				sheet.setDefaultRowHeightInPoints(70);
			}
			List<String> headers = listHeaders.get(m);

			for (int i = 0; i < headers.size(); i++) {
				HSSFCell cell = row.createCell(i);
				String header = headers.get(i);
				HSSFRichTextString text = new HSSFRichTextString(header);

				cell.setCellValue(text);
				// 单元格格式设置
				if (caseHeaders.contains(header)) {
					String titleColor = ExcelUtil.HEADER_COLOR_RECORD.get(header);
					if (titleColor != null && titleColor != "" && titleColor.equals("A")) {
						cell.setCellStyle(unalterTitleStyle);
					} else if (titleColor != null && titleColor != "" && titleColor.equals("B")) {
						cell.setCellStyle(updateTitleStyle);
					} else {
						cell.setCellStyle(castTitleStyle);
					}
				} else {
					String titleColor = ExcelUtil.HEADER_COLOR_RECORD.get(header);
					if (titleColor != null && titleColor != "" && titleColor.equals("C")) {
						cell.setCellStyle(castTitleStyle);
					} else {
						cell.setCellStyle(unalterTitleStyle);
					}
				}
			}
		}
		/** 设置行冻结，有几行标题，冻结几行 */
		sheet.createFreezePane(0, listHeaders.size(), 0, listHeaders.size());// 要冻结的列数，要冻结的行数，冻结列号，冻结行号
	}

	/**
	 * 插入表格的内容数据
	 * 
	 * @param listHeaders
	 *            复杂表头数据(结合合并单元格使用,实现复杂表头)
	 * @param datas
	 *            表格内容数据
	 * @param subheads
	 *            将表格数据设置为副表头; 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
	 * @param work
	 *            HSSFWorkbook
	 * @param sheet
	 *            HSSFSheet
	 * @param style2
	 *            表格格式
	 * @param rowHeight
	 *            设置行高; 参数1：哪一行 参数2：行高值
	 */
	private static void addData(List<List<String>> listHeaders, List<List<Object>> datas, List<Integer[]> subheads,
			HSSFWorkbook work, HSSFSheet sheet, HSSFCellStyle defaultStyle, Integer[] rowHeight,
			HSSFCellStyle passStyle, HSSFCellStyle failStyle) {
		List<String> headersFirst = listHeaders.get(0);
		List<String> headersSecond = listHeaders.size() > 1 ? listHeaders.get(1)
				: new ArrayList<String>(listHeaders.get(0).size());
		int headerSize = listHeaders.size();
		for (int rowNo = 0, max = datas.size(); rowNo < max; rowNo++) {
			List<Object> data = datas.get(rowNo);
			int rowNum = rowNo + headerSize;
			HSSFRow row = sheet.createRow((short) rowNum);// 计算出romNum
			if (rowHeight != null && rowHeight[0] == rowNum) {
				row.setHeightInPoints(rowHeight[1]);
			} else {
				// 设置默认行高
				sheet.setDefaultRowHeightInPoints(20);
			}
			for (int i = 0, j = data.size(); i < j; i++) {
				HSSFCell cell = row.createCell(i);
				GenerateXmlUtil.setValue(cell, data.get(i));
				String headerFirst = headersFirst.get(i);
				String headerSecond = headersSecond.get(i);
				boolean testResultCell = false;
				if ("P/F".equals(headerFirst) || "Pass/Failed".equals(headerFirst) || "P/F".equals(headerSecond)
						|| "Pass/Failed".equals(headerSecond)) {
					testResultCell = true;
				}
				if (testResultCell) {
					String value = data.get(i).toString();
					if ("Pass".equals(value)) {
						cell.setCellStyle(passStyle);
					} else if ("Failed".equals(value) || "Fail".equals(value)) {
						cell.setCellStyle(failStyle);
					} else {
						cell.setCellStyle(defaultStyle);
					}
				} else {
					cell.setCellStyle(defaultStyle);
				}
			}

		}
	}

	/**
	 * 设置表格格式
	 * 
	 * @param work
	 *            HSSFWorkbook
	 * @param fontName
	 *            字体样式
	 * @param size
	 *            字体大小
	 * @param bold
	 *            加粗
	 * @param center
	 *            居中
	 * @param foregroundColor
	 *            背景色
	 * @param wrapText
	 *            自动换行
	 * @param borderBottom
	 *            边框
	 * @return
	 */
	public static HSSFCellStyle setStyle(HSSFWorkbook work, String fontName, boolean center, String foregroundColor,
			boolean wrapText, boolean borderBottom, boolean fontColor) {
		// 创建字体样式
		HSSFFont font = work.createFont();
		// 创建格式
		HSSFCellStyle style = work.createCellStyle();
		// 设置字体
		if (fontName != null) {
			font.setFontName("宋体");
		}
		// 设置字体格式
		style.setFont(font);
		// 设置居中
		if (center) {
			// 水平居中
			style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		}
		// 设置背景颜色
		if (foregroundColor != null) {
			if ("黄色".equals(foregroundColor)) {
				style.setFillBackgroundColor(HSSFColor.YELLOW.index);
			} else if ("绿色".equals(foregroundColor)) {
				style.setFillForegroundColor(HSSFColor.CORNFLOWER_BLUE.index);// 前景填充色
			} else if ("红色".equals(foregroundColor)) {
				style.setFillForegroundColor(HSSFColor.RED.index);// 前景填充色

			} else if ("蓝色".equals(foregroundColor)) {
				style.setFillForegroundColor(HSSFColor.BLUE.index);// 前景填充色
				style.setFillBackgroundColor(HSSFColor.BLUE.index);
			} else if ("紫色".equals(foregroundColor)) {
				style.setFillForegroundColor(HSSFColor.VIOLET.index);// 前景填充色
			} else if ("灰色".equals(foregroundColor)) {
				style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);// 前景填充色
				style.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);// 背景
			}
		}
		// 设置边框
		if (borderBottom) {
			style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 下边框
			style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
			style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
			style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框
		}
		// 设置自动换行
		if (wrapText) {
			style.setWrapText(true);
		}
		if (fontColor) {
			font.setColor(HSSFColor.BLACK.index);
			style.setFont(font);
		}
		return style;
	}

	/**
	 * 最大列宽
	 * 
	 * @param headers
	 * @param datas
	 * @return
	 */
	private static List<Integer> getMaxCall(List<List<String>> listHeaders, List<List<Object>> datas) {
		// 创建最大列宽集合
		List<Integer> maxCall = new ArrayList<>();
		List<List<Integer>> lss = new ArrayList<>();
		// 计算标题行数据的列宽
		for (int i = 0, j = listHeaders.size(); i < j; i++) {
			List<Integer> hls = new ArrayList<>();
			for (int m = 0, n = listHeaders.get(i).size(); m < n; m++) {
				int length = listHeaders.get(i).get(m).getBytes().length;
				hls.add(length);
				if (i == 0) {
					// 最大列宽赋初值
					maxCall.add(0);
				}
			}
			lss.add(hls);
		}
		// 计算内容行数据的列宽
		for (int i = 0, j = datas.size(); i < j; i++) {
			List<Integer> dls = new ArrayList<>();
			for (int m = 0, n = datas.get(i).size(); m < n; m++) {
				Object obj = datas.get(i).get(m);
				if (obj.getClass() == Date.class) {
					// 日期格式类型转换
					obj = GenerateXmlUtil.DATE_FORAMT.format(obj);
				}
				int length = obj.toString().getBytes().length;
				dls.add(length);
			}
			lss.add(dls);
		}
		// // 根据列宽计算出每列的最大宽度
		// for (int i = 0, j = lss.size(); i < j; i++) {
		// for (int m = 0, n = lss.get(i).size(); m < n; m++) {
		// Integer a = lss.get(i).get(m);
		// Integer b = maxCall.get(m);
		// if (a > b) {
		// maxCall.set(m, a);
		// }
		// }
		// }
		return maxCall;
	}

	/**
	 * 表格数据类型转换
	 * 
	 * @param cell
	 * @param value
	 */
	private static void setValue(Cell cell, Object value) {
		if (value == null) {
			cell.setCellValue("");
		} else if (value.getClass() == String.class) {
			cell.setCellValue((String) value);
		} else if (value.getClass() == Integer.class) {
			cell.setCellValue((Integer) value);
		} else if (value.getClass() == Double.class) {
			cell.setCellValue((Double) value);
		} else if (value.getClass() == Date.class) {
			cell.setCellValue(GenerateXmlUtil.DATE_FORAMT.format(value));
		} else if (value.getClass() == Long.class) {
			cell.setCellValue((Long) value);
		}
	}

	/**
	 * Description 添加表格有效值
	 * 
	 * @param sheet
	 * @param listHeaders
	 * @param dataSize
	 */
	public static void addValidation(HSSFSheet sheet, List<List<String>> listHeaders, Integer dataSize) {
		int firstRow = listHeaders.size();
		int endRow = firstRow + (dataSize == 0 ? 1 : dataSize);
		for (List<String> headers : listHeaders) {
			for (String header : headers) {
				String field = ExcelUtil.HEADER_MAP.get(header);
				if (field == null || "".equals(field)) {
					continue;
				}
				int indexCol = headers.indexOf(header);
				if ("Category".equals(field)) {
					setHSSFValidation(sheet,
							ExcelUtil.CURRENT_CATEGORIES.toArray(new String[ExcelUtil.CURRENT_CATEGORIES.size()]),
							firstRow, endRow, indexCol, indexCol);
				} else {
					String fieldType = ExcelUtil.FIELD_TYPE_RECORD.get(field);
					if ("pick".equals(fieldType)) {
						List<String> valueList = ExcelUtil.PICK_FIELD_RECORD.get(field);

						setHSSFValidation(sheet, valueList.toArray(new String[valueList.size()]), firstRow, endRow,
								indexCol, indexCol);
					}
				}
			}
		}
	}

	/**
	 * Description 设置有效值验证，有效值验证不能超过255个字符
	 * 
	 * @param valueList
	 * @param firstRow
	 * @param endRow
	 * @param FirstCol
	 * @param endCod
	 */
	public static void setHSSFValidation(HSSFSheet sheet, String[] valueList, int firstRow, int endRow, int FirstCol,
			int endCod) {
		DVConstraint constraint = DVConstraint.createExplicitListConstraint(valueList);
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, FirstCol, endCod);
		HSSFDataValidation validationList = new HSSFDataValidation(regions, constraint);
		sheet.addValidationData(validationList);
	}

}