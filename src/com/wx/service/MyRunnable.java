package com.wx.service;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import com.wx.ui.TestResultExportUI;
import com.wx.util.Constants;
import com.wx.util.MKSCommand;
@SuppressWarnings("all")
public class MyRunnable implements Runnable {

	public MKSCommand cmd;
	public List<String> tsIds = new ArrayList<>();//存放获取到的ID集合
	public String filePath;
	public static String documentName;
	public MyRunnable() {
		super();
	}
	
	@Override
	public void run() {
		boolean success = true;
		try {
//			String userHome = TestObjReportUI.ENVIRONMENTVAR.get(Constants.USER_HOME); //获取用户目录
//			if (userHome == null || userHome.length() == 0) {
//				userHome = TestObjReportUI.ENVIRONMENTVAR.get("USERPROFILE");//如果没获取到 手动再次获取
//			}
//			if (userHome == null || userHome.length() == 0) {
//				userHome = "C:\\Users\\" + TestObjReportUI.ENVIRONMENTVAR.get(Constants.USERNAME);//再次获取用户目录
//			}
//			filePath = input.toString();
//			if (!input.exists()) { //如果不存在就去创建
//				input.mkdirs();
//			}
			
			TestResultExportUI.logger.info("GET MKS connection completed!"); 
			TestResultExportUI.logger.info("Check the document ID completed!"); 
			ExcelUtil util = new ExcelUtil();
			TestResultExportUI.logger.info("start to export Test Suite report!");
			
			String filePath       = TestResultExportUI.class.newInstance().filePath; 
		    util.exportReport(tsIds, cmd, filePath);
			
		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestResultExportUI.contentPane, e.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
			TestResultExportUI.logger.info("Error: " + e.getMessage());
			e.printStackTrace();
			success = false;
		} finally {
			if(success) {
				JOptionPane.showMessageDialog(TestResultExportUI.contentPane, "Success!");
			}
			try {
				cmd.release();
				TestResultExportUI.logger.info("cmd release!");
			} catch (IOException e) {
				TestResultExportUI.logger.info("Error: " + e.getMessage());
			} finally {
				System.exit(0);
			}
		}
	}
}
