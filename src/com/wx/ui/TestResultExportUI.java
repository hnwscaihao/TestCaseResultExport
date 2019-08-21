package com.wx.ui;

import java.awt.EventQueue;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.wx.service.ExcelUtil;
import com.wx.service.MyRunnable;
import com.wx.util.Constants;
import com.wx.util.MKSCommand;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;

import java.awt.Label;
import java.awt.Font;
import java.awt.Color;
import java.awt.Button;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.ActionEvent;
import javax.swing.JLabel;
import javax.swing.border.EtchedBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.JButton;
import javax.swing.SwingConstants;
import javax.swing.JTextField;
@SuppressWarnings("all")
public class TestResultExportUI extends JFrame {
	private static final long serialVersionUID = 1L;
	public static JPanel contentPane;
	public static final Map<String, String> ENVIRONMENTVAR = System.getenv();
	public static MKSCommand cmd;
	public static List<String> tsIds = new ArrayList<>();
	public static final Logger logger = Logger.getLogger(TestResultExportUI.class);
	JProgressBar progressBar;
	private static String DOCUMENT_TYPE ; 
	private boolean start;
	private MyRunnable run = new MyRunnable();
	private Thread thread = new Thread();// 查询线程
	private JButton button;
	private JLabel label_1;
	public  static int MIN_PROGRESS = 0;
	public  static int MAX_PROGRESS = 8;
	public static int currentProgress = MIN_PROGRESS;
	public static ScheduledExecutorService pool;
	public static boolean isParseSuccess = false;//判断导出是否结束
	public static  String filePath;
	public static boolean ispause=false;
	public static boolean isStart=false;
	public static boolean chooseRight = false;
	private ExcelUtil excelUtil = new ExcelUtil();
	
	private static String documentName ;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					initMksCommand();//初始化MKSCommand中的参数，并获得连接
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());//设置与本机适配的swing样式
					TestResultExportUI frame = new TestResultExportUI();
					frame.setVisible(true);
					frame.setLocationRelativeTo(null);
					getSelectedIdList();//获取到当前选中的id添加进集合Ids集合
				} catch (Exception e) {
					JOptionPane.showMessageDialog(contentPane, e.getMessage());
					System.exit(0);
				}
			}
		});
	}

	/**
	 * 开始导出
	 * @throws Exception
	 */
	public void startExport() throws Exception {
		run.documentName = documentName;
		run.tsIds = tsIds;
		run.cmd = cmd;
		thread = new Thread(run);
		thread.start();//
	}
	
	
	public void  updateProgress (){
	    pool = Executors.newScheduledThreadPool(1);
	    pool.scheduleAtFixedRate(new Runnable() {
			@Override
			public void run() {
				currentProgress++;
			        if (currentProgress > MAX_PROGRESS && !isParseSuccess) {
			            currentProgress = MIN_PROGRESS;
			        }
			        if(isParseSuccess) {
			        	getProgressBar().setValue(MAX_PROGRESS);
			        	pool.shutdown();
			        }
			        getProgressBar().setValue(currentProgress);
			}
			}, 
			1, 
			1,
			TimeUnit.SECONDS);
	}
	
		
	
	
	/**
	 * Create the frame.
	 */
	public TestResultExportUI() {
		setAutoRequestFocus(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle("Export Test Case & Result");
		setResizable(false);
		setBounds(100, 100, 714, 252);
		contentPane = new JPanel();
		contentPane.setForeground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		progressBar = new JProgressBar();
		progressBar.setString("Exporting, please wait!");
		progressBar.setBounds(82, 133, 361, 24);
        // 设置当前进度值
        getProgressBar().setValue(0);
        getProgressBar().setMinimum(MIN_PROGRESS);
        getProgressBar().setMaximum(MAX_PROGRESS);
		contentPane.add(progressBar);
		
		button = new JButton("Export");
		
		button.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				if (label_1.getText().equals("< Please select an export path />")) {
					
					JOptionPane.showConfirmDialog(contentPane, "Please select a folder as the path to export!!");//导出路径
					return;
				}else{
					updateProgress();
				}
				//开始导出！
				try {
					startExport();
					button.setEnabled(false); 
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(contentPane, e1.getMessage());
					System.exit(0);
				}
			}
		});
		button.setForeground(Color.BLACK);
		button.setBounds(526, 127, 137, 30);
		contentPane.add(button);
		label_1 = new JLabel("< Please select an export path />");
		label_1.setForeground(Color.BLACK);
		label_1.setBorder(new EtchedBorder());
		label_1.setBounds(82, 52, 361, 37);
		contentPane.add(label_1);
		
		JButton button_1 = new JButton("Browse");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				TestResultExportUI.this.browseDocAction();//
			}
		});
		button_1.setBounds(526, 56, 137, 29);
		contentPane.add(button_1);
		
		JLabel label = new JLabel("*");
		label.setForeground(Color.RED);
		label.setBounds(56, 60, 24, 21);
		contentPane.add(label);

	}
	/**
	 * 浏览文件操作
	 */
	protected void browseDocAction() {
		logger.info("Start to load file of import"); 
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle("选择一个路径");
		fc.setAcceptAllFileFilterUsed(true);
		fc.setMultiSelectionEnabled(false);
		fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
		int returnVal = fc.showOpenDialog(this);
		if (returnVal == 0) {
			File input = fc.getSelectedFile();
			filePath = input.toString();
			logger.info("selected file:" + input);
			if(!input.isDirectory() && !input.toString().endsWith("xls") ){//生成03版本，所以是xls后缀
				JOptionPane.showMessageDialog(contentPane, "Please Choose a folder or a XLS file");
				chooseRight = false;
			}else{
				label_1.setText("path:"+input.toString());
				chooseRight = true;
			}
		}
	}
	
	public JProgressBar getProgressBar() {
		return progressBar;
	}


	public void setProgressBar(JProgressBar progressBar) {
		this.progressBar = progressBar;
	}

	/**
	 * 初始化MKSCommand中的参数，并获得连接
	 */
	public static void initMksCommand() {
		try {
			String host = TestResultExportUI.ENVIRONMENTVAR.get(Constants.MKSSI_HOST);
			if(host==null || host.length()==0) {
				host = "172.25.4.18";
			}
			cmd = new MKSCommand(host, 7001, "admin", "admin", 4, 16);
//			cmd.getSession();
		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestResultExportUI.contentPane, "Can not get a connection!", "Message",
					JOptionPane.WARNING_MESSAGE);
			TestResultExportUI.logger.info("Can not get a connection!");
			System.exit(0);
		}
	}

	/**
	 * 获取当前选中id的List集合
	 * @return
	 * @throws Exception
	 */
	public static List<String> getSelectedIdList() throws Exception {
		String issueCount = TestResultExportUI.ENVIRONMENTVAR.get(Constants.MKSSI_NISSUE);
		TestResultExportUI.logger.info("get issue count from environment:" + issueCount); 
		if (issueCount != null && issueCount.trim().length() > 0) { 
			for (int index = 0; index < Integer.parseInt(issueCount); index++) {
				String id = TestResultExportUI.ENVIRONMENTVAR.get(String.format(Constants.MKSSI_ISSUE_X, index));
				TestResultExportUI.logger.info("get the selection test Suite : " + id);
				tsIds.add(id);//获取到当前选中的id添加进集合Ids集合
			}
		} else {
			TestResultExportUI.logger.info("No ID was obtained!!! :" + issueCount); 
		}
		if (tsIds.size() > 0) {//如果选中的id集合不为空，通过id获取条目简要信息
			List<Map<String, String>> itemByIds = cmd.getItemByIds(tsIds, Arrays.asList("ID", "Type","Document Short Title"));
			List<String> notTSList = new ArrayList<>();
			for (Map<String, String> map : itemByIds) {
				DOCUMENT_TYPE = map.get("Type");
				String id = map.get("ID");
				documentName = map.get("Document Short Title");
				if(!ExcelUtil.TEST_SUITE.equals(DOCUMENT_TYPE)){
					JOptionPane.showConfirmDialog(contentPane, "Please Select the " + ExcelUtil.TEST_SUITE + " to Export!");//导出路径
					System.exit(0);
					return null;
				}
			}
			if (notTSList.size() > 0) {
//				throw new Exception("This item " + notTSList + " is not [ " + documentType + " ]! Please  select the right type!");
			} else {
				TestResultExportUI.logger.info("get the selection Document : " + tsIds);
			}
		} else {
			throw new Exception("Please select the ID of a Document!");
		}
		return tsIds;
	}
}
