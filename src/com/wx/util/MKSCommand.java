package com.wx.util;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.log4j.Logger;

import com.mks.api.CmdRunner;
import com.mks.api.Command;
import com.mks.api.IntegrationPoint;
import com.mks.api.IntegrationPointFactory;
import com.mks.api.MultiValue;
import com.mks.api.Option;
import com.mks.api.SelectionList;
import com.mks.api.Session;
import com.mks.api.response.APIException;
import com.mks.api.response.Field;
import com.mks.api.response.Item;
import com.mks.api.response.ItemList;
import com.mks.api.response.Response;
import com.mks.api.response.WorkItem;
import com.mks.api.response.WorkItemIterator;
import com.wx.service.ExcelUtil;
import com.wx.ui.TestResultExportUI;

public class MKSCommand {

	private static final Logger logger = Logger.getLogger(MKSCommand.class.getName());
	private Session mksSession = null;
	private IntegrationPointFactory mksIpf = null;
	private IntegrationPoint mksIp = null;
	private static CmdRunner mksCmdRunner = null;
	private Command mksCommand = null;
	private Response mksResponse = null;
	private boolean success = false;
	private String currentCommand;
	private String hostname = null;
	private int port = 7001;
	private String user;
	private String password;
	private int APIMajor = 4;
	private int APIMinor = 16;
	private static String errorLog;
	private static final String FIELDS = "fields";
	private static final String CONTAINS = "Contains";
	
	private static final SimpleDateFormat FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
	public MKSCommand(String _hostname, int _port, String _user, String _password, int _apimajor, int _apiminor) {
		hostname = _hostname;
		port = _port;
		user = _user;
		password = _password;
//		createSession();
		getSession();
	}

	public MKSCommand(String args[]) {
		hostname = args[0];
		port = Integer.parseInt(args[1]);
		user = args[2];
		password = args[3];
		APIMajor = Integer.parseInt(args[4]);
		APIMinor = Integer.parseInt(args[5]);
		createSession();
	}

	public void setCmd(String _type, String _cmd, ArrayList<Option> _ops, String _sel) {
		mksCommand = new Command(_type, _cmd);
		String cmdStrg = (new StringBuilder(String.valueOf(_type))).append(" ").append(_cmd).append(" ").toString();
		if (_ops != null && _ops.size() > 0) {
			for (int i = 0; i < _ops.size(); i++) {
				cmdStrg = (new StringBuilder(String.valueOf(cmdStrg))).append(_ops.get(i).toString()).append(" ")
						.toString();
				// Option o = new Option(_ops.get(i).toString());
				mksCommand.addOption(_ops.get(i));
			}

		}
		if (_sel != null && _sel != "") {
			cmdStrg = (new StringBuilder(String.valueOf(cmdStrg))).append(_sel).toString();
			mksCommand.addSelection(_sel);
		}
		currentCommand = cmdStrg;
		// logger.info((new StringBuilder("Command:
		// ")).append(cmdStrg).toString());
	}

	public String getCommandAsString() {
		return currentCommand;
	}

	public boolean getResultStatus() {
		return success;
	}

	public String getConnectionString() {
		String c = (new StringBuilder(String.valueOf(hostname))).append(" ").append(port).append(" ").append(user)
				.append(" ").append(password).toString();
		return c;
	}

	public void exec() {
		success = false;
		try {
			mksResponse = mksCmdRunner.execute(mksCommand);
			// logger.info((new StringBuilder("Exit Code:
			// ")).append(mksResponse.getExitCode()).toString());
			success = true;
		} catch (APIException ae) {
			logger.error(ae.getMessage());
			success = false;
			errorLog = ae.getMessage();
		} catch (NullPointerException npe) {
			success = false;
			logger.error(npe.getMessage(), npe);
			errorLog = npe.getMessage();
		}
	}

	public void release() throws IOException {
		try {
			if (mksSession != null) {
				mksCmdRunner.release();
				mksSession.release();
				mksIp.release();
				mksIpf.removeIntegrationPoint(mksIp);
			}
			success = false;
			currentCommand = "";
		} catch (APIException ae) {
			logger.error(ae.getMessage(), ae);
		}
	}

	public void getSession() {
		try {
			mksIpf = IntegrationPointFactory.getInstance();
			mksIp = mksIpf.createLocalIntegrationPoint(APIMajor, APIMinor);
			mksIp.setAutoStartIntegrityClient(true);
			mksSession = mksIp.getCommonSession();
			mksCmdRunner = mksSession.createCmdRunner();
			mksCmdRunner.setDefaultUsername(user);
			mksCmdRunner.setDefaultPassword(password);
			mksCmdRunner.setDefaultHostname(hostname);
			mksCmdRunner.setDefaultPort(port);
		} catch (APIException ae) {
			logger.error(ae.toString(), ae);
		}
	}

	@SuppressWarnings("deprecation")
	public void createSession() {
		try {
			mksIpf = IntegrationPointFactory.getInstance();
			mksIp = mksIpf.createIntegrationPoint(hostname, port, APIMajor, APIMinor);
			mksSession = mksIp.createSession(user, password);
			mksCmdRunner = mksSession.createCmdRunner();
			mksCmdRunner.setDefaultHostname(hostname);
			mksCmdRunner.setDefaultPort(port);
			mksCmdRunner.setDefaultUsername(user);
			mksCmdRunner.setDefaultPassword(password);
		} catch (APIException ae) {
			logger.error(ae.getMessage(), ae);
		}
	}

	public String[] getResult() {
		String result[] = null;
		int counter = 0;
		try {
			WorkItemIterator mksWii = mksResponse.getWorkItems();
			result = new String[mksResponse.getWorkItemListSize()];
			while (mksWii.hasNext()) {
				WorkItem mksWi = mksWii.next();
				Field mksField;
				for (Iterator<?> mksFields = mksWi.getFields(); mksFields.hasNext();) {
					mksField = (Field) mksFields.next();
					result[counter] = mksField.getValueAsString();
				}

				counter++;
			}
		} catch (APIException ae) {
			logger.error(ae.toString(), ae);
			JOptionPane.showMessageDialog(null, ae.toString(), "ERROR", 0);
		} catch (NullPointerException npe) {
			logger.error(npe.toString(), npe);
			JOptionPane.showMessageDialog(null, npe.toString(), "ERROR", 0);
		}
		return result;
	}

	/**
	 * 根据Ids查询字段的值
	 * 
	 * @param ids
	 * @param fields
	 * @return
	 * @throws APIException
	 */
	public List<Map<String, String>> getItemByIds(List<String> ids, List<String> fields) throws APIException {
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		Command cmd = new Command("im", "issues");
		MultiValue mv = new MultiValue();
		mv.setSeparator(",");
		for (String field : fields) {
			mv.add(field);
		}
		Option op = new Option("fields", mv);
		cmd.addOption(op);

		SelectionList sl = new SelectionList();
		for (String id : ids) { 
			String splitID =null;
			if(id.startsWith("[")&&id.endsWith("]")){
				splitID = id.substring(id.indexOf("[")+1,id.indexOf("]"));
				sl.add(splitID.trim());
			}else if(id.startsWith("[")){
				splitID = id.substring(id.indexOf("[")+1,id.length());
				sl.add(splitID.trim());
			}else if(id.endsWith("]")){
				splitID = id.substring(0,id.indexOf("]"));
				sl.add(splitID.trim());
			}else if(id.startsWith(" ")){
				splitID =id.substring(1,id.length());
				sl.add(splitID.trim());
			}else{
				sl.add(id.trim());
			}
		}
		cmd.setSelectionList(sl);

		Response res = null;
		try {
			res = mksCmdRunner.execute(cmd);
			WorkItemIterator it = res.getWorkItems();
			while (it.hasNext()) {
				WorkItem wi = it.next();
				Map<String, String> map = new HashMap<String, String>();
				for (String field : fields) {
					if (field.contains("::")) {
						field = field.split("::")[0];
					}
					String value = wi.getField(field).getValueAsString(); 
					map.put(field, value);
				}
				list.add(map);
			}
		} catch (APIException e) {
			// success = false;
			logger.error(e.getMessage());
			throw e;
		}
		return list;
	}

	public boolean getResultState() {
		return success;
	}

	public String getErrorLog() {
		return errorLog;
	}

	
	@Deprecated
	public List<Map<String, String>>  getAllChild(List<String> ids, List<String> childs) throws APIException {
		List<Map<String, String>> itemByIds = getItemByIds(ids, Arrays.asList("ID", "Contains"));//查询文档id包含字段heading
		for(Map<String,String> map : itemByIds) { //
			String contains = map.get("Contains");
			String id = map.get("ID");
			map.put("ID", id);
			if(contains!=null && contains.length()>0) {
//				List<String> childIds = Arrays.asList(contains.replaceAll("ay", "").split(","));
				getAllChild(Arrays.asList(id), Arrays.asList(contains));
			}
		}
		return itemByIds;
		
	}
	
	public SelectionList contains(SelectionList documents) throws APIException {
		return relationshipValues(CONTAINS, documents);
	}

	public SelectionList relationshipValues(String fieldName, SelectionList ids) throws APIException {
		if (fieldName == null) {
			throw new APIException("invoke fieldValues() ----- fieldName is null.");
		}
		if (ids == null || ids.size() < 1) {
			throw new APIException("invoke fieldValues() ----- ids is null or empty.");
		}
		Command command = new Command(Command.IM, Constants.ISSUES);
		command.addOption(new Option(Constants.FIELDS, fieldName));
		command.setSelectionList(ids);
		Response res = mksCmdRunner.execute(command);
		WorkItemIterator it = res.getWorkItems();
		SelectionList contents = new SelectionList();
		while (it.hasNext()) {
			WorkItem wi = it.next();
			ItemList il = (ItemList) wi.getField(fieldName).getList();
			if(il != null) {
				for (int i = 0; i < il.size(); i++) {
					Item item = (Item) il.get(i);
					String id = item.getId();
					contents.add(id);
				}
			}
		}
		return contents;
	}
	
	public List<Map<String, String>> allContents(String document, List<String> fieldList) throws APIException ,Exception {
		List<Map<String, String>> returnResult = new ArrayList<Map<String, String>>();
		Command command = new Command("im", "issues");
		command.addOption(new Option(FIELDS, CONTAINS));
		command.addSelection(document);
		Response res = mksCmdRunner.execute(command);
		WorkItemIterator it = res.getWorkItems();
		SelectionList sl = new SelectionList();
		List<String> fields = new ArrayList<>();
		fields.add("ID");
		if (fieldList != null) {
			fields.addAll(fieldList);
		}
		while (it.hasNext()) {
			WorkItem wi = it.next();
			ItemList il = (ItemList) wi.getField(CONTAINS).getList();
			for (int i = 0; i < il.size(); i++) {
				Item item = (Item) il.get(i);
				String id = item.getId();
				sl.add(id);
			}
		}
		SelectionList contents = null;
		if (sl != null && sl.size() >= 1) {
			contents = contains(sl);

			if (contents.size() > 0) {
				SelectionList contains = new SelectionList();
				contains.add(contents);
				while (true) {
					SelectionList conteins = contains(contains);
					if (conteins.size() < 1) {
						break;
					}
					contents.add(conteins);
					contains = new SelectionList();
					contains.add(conteins);
				}
			}
			contents.add(sl);
			List<Map<String, String>> list = new ArrayList<Map<String, String>>();
			if (contents.size() > 500) {
				List<SelectionList> parallel = new ArrayList<SelectionList>();
				SelectionList ids = new SelectionList();
				for (int i = 0;; i++) {
					if (i % 500 == 0 && ids.size() > 0) {
						parallel.add(ids);
						ids = new SelectionList();
					}
					ids.add(contents.getSelection(i));
					if (i + 1 == contents.size()) {
						parallel.add(ids);
						break;
					}
				}
				for (SelectionList selectionList : parallel) {
					list.addAll(queryIssues(selectionList, fields));
				}
			} else {
				list.addAll(queryIssues(contents, fields));
			}
		}
		return returnResult;
	}
	
	public List<List<Map<String, String>>> allContentsByHeading(String document, List<String> fieldList) throws APIException ,Exception {
		List<List<Map<String, String>>> returnResult = new ArrayList<List<Map<String, String>>>();
		List<String> firstContainIds = new ArrayList<>();//第一级ID
		Command command = new Command("im", "issues");
		command.addOption(new Option(FIELDS, CONTAINS));
		command.addSelection(document);
		Response res = mksCmdRunner.execute(command);
		WorkItemIterator it = res.getWorkItems();
		SelectionList sl = new SelectionList();
		List<String> fields = new ArrayList<>();
		fields.add("ID");
		if (fieldList != null) {
			fields.addAll(fieldList);
		}
		while (it.hasNext()) {
			WorkItem wi = it.next();
			ItemList il = (ItemList) wi.getField(CONTAINS).getList();
			for (int i = 0; i < il.size(); i++) {
				Item item = (Item) il.get(i);
				String id = item.getId();
				firstContainIds.add(id);
			}
		}
		for(String id : firstContainIds){
			sl = new SelectionList();
			SelectionList contents = null;
			sl.add(id);
			if (sl != null && sl.size() >= 1) {
				contents = contains(sl);
				
				if (contents.size() > 0) {
					SelectionList contains = new SelectionList();
					contains.add(contents);
					while (true) {
						SelectionList conteins = contains(contains);
						if (conteins.size() < 1) {
							break;
						}
						contents.add(conteins);
						contains = new SelectionList();
						contains.add(conteins);
					}
				}
				sl.add(contents);
				List<Map<String, String>> list = new ArrayList<Map<String, String>>();
				if (sl.size() > 500) {
					List<SelectionList> parallel = new ArrayList<SelectionList>();
					SelectionList ids = new SelectionList();
					for (int i = 0;; i++) {
						if (i % 500 == 0 && ids.size() > 0) {
							parallel.add(ids);
							ids = new SelectionList();
						}
						ids.add(sl.getSelection(i));
						if (i + 1 == sl.size()) {
							parallel.add(ids);
							break;
						}
					}
					for (SelectionList selectionList : parallel) {
						list.addAll(queryIssues(selectionList, fields));
					}
				} else {
					list.addAll(queryIssues(sl, fields));
				}
				returnResult.add(list);
			}
		}
		return returnResult;
	}
	
	public List<Map<String, String>> queryIssues(SelectionList selectionList,List<String> fields) throws APIException, Exception {
		List<Map<String, String>> returnResult = new ArrayList<Map<String,String>>();
		boolean needFilter = false;
		String category = "";
		fields.add("Category");
		Command cmd = new Command("im", "issues");
		MultiValue mv = new MultiValue();
		mv.setSeparator(",");
		for (String field : fields) {
			mv.add(field);
		}
		Option op = new Option("fields", mv);
		cmd.addOption(op);
		cmd.setSelectionList(selectionList);
		Response res = null;
		try {
			res = mksCmdRunner.execute(cmd);
			WorkItemIterator it = res.getWorkItems();
			while (it.hasNext()) {
				WorkItem wi = it.next();
				Map<String, String> map = new HashMap<String, String>();
				for (String field : fields) {
					if (field.contains("::")) {
						field = field.split("::")[0];
					}
					Field fieldObj = wi.getField(field);
					String fieldType = fieldObj.getDataType();
					String value = fieldObj.getValueAsString()!=null?fieldObj.getValueAsString().toString():null;
					value = parseDateVal(value, fieldType);
					if(value !=null && value.contains("[") && value.contains("]")){
						value = value.substring(value.indexOf("[")+1, value.indexOf("]"));
					}
					if("[]".equals(value)){
						value = null;
					}
					map.put(field, value);
				}
				boolean canAdd = true;
				if(needFilter){
					String currentCategory = map.get("Category");
					if(!currentCategory.equals(category))
						canAdd = false;
				}
				if(canAdd)
					returnResult.add(map);
			}
		} catch (APIException e) {
			logger.error(e.getMessage());
			throw e;
		}
		return returnResult;
	}
	
	public static String parseDateVal(String value, String fieldType){
		if("java.util.Date".equals(fieldType)){
			value = FORMAT.format(new Date(value));
		}
		return value;
	}
	
	/**
	 * 查询系统所有用户信息
	 * @return
	 * @throws APIException
	 */
	public void getAllUsers() throws APIException{
		String userName = null;
		String loginId = null;
		Command cmd = new Command(Command.IM, "users");
		cmd.addOption(new Option("fields", "name,fullname,email,isActive"));
		Response res = mksCmdRunner.execute(cmd);
		if (res != null) {
			WorkItemIterator iterator = res.getWorkItems();
			while (iterator.hasNext()) {
				WorkItem item = iterator.next();
				if (item.getField("isActive").getValueAsString().equalsIgnoreCase("true")) {
					userName = item.getField("fullname").getValueAsString();
					loginId = item.getField("name").getValueAsString();
					ExcelUtil.USER_MAP.put(loginId, userName);
				}
			}
		}
	}
	
	public List<String> getTestStepFields() throws APIException {
		List<String> fieldList = new ArrayList<>();
		if (fieldList.isEmpty()) {
			fieldList.add("ID");
			fieldList.add("Description");
		}
		return fieldList;
	}
	
	public static List<Map<String, String>> findIssuesByQueryDef(List<String> fields, String query) throws APIException {
		if (query == null || query.isEmpty()) {
			throw new APIException("invoke findIssuesByQueryDef() ----- query is null or empty.");
		}
		if (fields == null) {
			fields = new ArrayList<>();
		}
		if (fields.size() < 1) {
			fields.add("ID");
			fields.add("Project");
			fields.add("Type");
			fields.add("State");
		}
		MultiValue mv = new MultiValue(",");
		for (String field : fields) {
			mv.add(field);
		}
		Command command = new Command(Command.IM, Constants.ISSUES);
		command.addOption(new Option(Constants.FIELDS, mv));
		command.addOption(new Option(Constants.QUERY_DEFINITION, query));
//		command.addOption(new Option("showTestResults"));
		Response res = mksCmdRunner.execute(command);
		WorkItemIterator it = res.getWorkItems();
		List<Map<String, String>> list = new ArrayList<>();
		while (it.hasNext()) {
				WorkItem wi = it.next();
				Iterator<?> iterator = wi.getFields();
				Map<String, String> map = new HashMap<>();
				while (iterator.hasNext()) {
					Field field = (Field) iterator.next();
					String fieldName = field.getName();
					if (Constants.ITEMLIST.equals(field.getDataType())) {
						StringBuilder sb = new StringBuilder();
						ItemList il = (ItemList) field.getList();
						for (int i = 0; i < il.size(); i++) {
							Item item = (Item) il.get(i);
							if (i > 0) {
								sb.append(",");
							}
							sb.append(item.getId());
						}
						map.put(fieldName, sb.toString());
					} else {
						map.put(fieldName, field.getValueAsString());
					}
				}
				list.add(map);
				
			}
		return list;
	}
				
	public void editIssue(String id, Map<String, String> fieldValue, Map<String, String> richFieldValue)
			throws APIException {
		Command cmd = new Command(Command.IM, "editissue");
		if (fieldValue != null) {
			for (Map.Entry<String, String> entrty : fieldValue.entrySet()) {
				cmd.addOption(new Option("field", entrty.getKey() + "=" + entrty.getValue()));
			}
		}
		if (richFieldValue != null) {
			for (Map.Entry<String, String> entrty : richFieldValue.entrySet()) {
				cmd.addOption(new Option("richContentField", entrty.getKey() + "=" + entrty.getValue()));
			}
		}
		cmd.addSelection(id);
		mksCmdRunner.execute(cmd);
	}
	
	public List<String> viewIssue(String id, boolean showRelationship)
			throws APIException {
		Command cmd = new Command(Command.IM, "viewissue");
		MultiValue mv = new MultiValue(",");
		cmd.addOption(new Option("showTestResults"));
		if(showRelationship){
			cmd.addOption(new Option("showRelationships"));
		}
		cmd.addSelection(id);
		Response res = mksCmdRunner.execute(cmd);
		WorkItemIterator it = res.getWorkItems();
		List<String> relations = new ArrayList<String>();
		while (it.hasNext()) {
			WorkItem wi = it.next();
			Iterator<?> iterator = wi.getFields();
			Map<String, String> map = new HashMap<>();
			while (iterator.hasNext()) {
				Field field = (Field) iterator.next();
				String fieldName = field.getName();
//				if("MKSIssueTestResults".equals(fieldName)){
//					field.getList();
//				}
				if("Test Steps".equals(fieldName)){
					System.out.println("123");
					StringBuilder sb = new StringBuilder();
					ItemList il = (ItemList) field.getList();
					for (int i = 0; i < il.size(); i++) {
						Item item = (Item) il.get(i);
						if (i > 0) {
							sb.append(",");
						}
						sb.append(item.getId());
					}
					map.put(fieldName, sb.toString());
				}
				if("Test Result".equals(fieldName) || "Test Results".equals(fieldName)){
					System.out.println("123");
				}
			}
		}
		return relations;
	}

	public List<Map<String, Object>> getResult(String sessionID, String suiteID, String type) throws APIException {
		List<Map<String, Object>> result = new ArrayList<>();
		Command cmd = new Command("tm", "results");
		
		cmd.addOption(new Option("caseID", suiteID));
		List<String> fields = new ArrayList<>();
		fields.add("sessionID");
		fields.add("verdict"); 
		fields.add("Annotation"); 
		
		MultiValue mv = new MultiValue();
		mv.setSeparator(",");
		for (String field : fields) {
			mv.add(field);
		}
		Option op = new Option("fields", mv);
		cmd.addOption(op);
		Response res = null;
		if (type.equals("Test Suite")) {
			res = mksCmdRunner.execute(cmd);
			WorkItemIterator wk = res.getWorkItems();
			while (wk.hasNext()) {
				Map<String, Object> map = new HashMap<>();
				WorkItem wi = wk.next();
				for (String field : fields) {
					Object value = wi.getField(field).getValue();
					map.put(field, value);
				}
				result.add(map);
			}
		} else if (type.equals("Test Case")) {
			try {
				res = mksCmdRunner.execute(cmd);
				WorkItemIterator wk = res.getWorkItems();
				while (wk.hasNext()) {
					Map<String, Object> map = new HashMap<>();
					WorkItem wi = wk.next();
					for (String field : fields) {
						Object value = wi.getField(field).getValue();
						if(value instanceof Item){
							Item item = (Item) value;
							value = item.getId();
						}
						map.put(field, value);
					}
					result.add(map);
				}
			} catch (Exception e) {
				e.printStackTrace();
				
			}
		}
		return result;
	}
	
	/**
	 * Description 获取所有Field 类型，并把Pick值预先取出
	 * @param fields
	 * @param PICK_FIELD_RECORD
	 * @return
	 * @throws APIException
	 */
	public Map<String,String> getAllFieldType(List<String> fields, Map<String,List<String>> PICK_FIELD_RECORD) throws APIException{
		Map<String,String> fieldTypeMap = new HashMap<String,String>();
		Command cmd = new Command("im", "fields");
		cmd.addOption(new Option("noAsAdmin"));
		cmd.addOption(new Option("fields", "picks,type"));
		for(String field : fields){
			if(field!=null && field.length()>0){
				cmd.addSelection(field);
			}
		}
		Response res=null;
		try {
			res = mksCmdRunner.execute(cmd);
		} catch (APIException e) {
			
			e.printStackTrace();
			System.out.println(e.getMessage());
		}
		
		if (res != null) {
			WorkItemIterator it = res.getWorkItems();
			while (it.hasNext()) {
				WorkItem wi = it.next();
				String field = wi.getId();
				String fieldType = wi.getField("Type").getValueAsString();
				if("pick".equals(fieldType) ){
					Field picks = wi.getField("picks");
					ItemList itemList = (ItemList) picks.getList();
					if (itemList != null) {
						List<String> pickVals = new ArrayList<String>();
						for (int i = 0; i < itemList.size(); i++) {
							Item item = (Item) itemList.get(i);
							String visiblePick = item.getId();
							Field attribute = item.getField("active");
							if (attribute != null && attribute.getValueAsString().equalsIgnoreCase("true")
									&& !pickVals.contains(visiblePick)) {
								pickVals.add(visiblePick);
							}
						}
						PICK_FIELD_RECORD.put(field, pickVals);
					}
				}else if("fva".equals(fieldType)){
					
				}
				fieldTypeMap.put(field, fieldType);
			}
		}
		return fieldTypeMap;
	}
	
	/**
	 * Description 查询所有Projects
	 * @return
	 * @throws APIException
	 */
	public List<String> getProjects() throws APIException{
		List<String> projects = new ArrayList<String>();
		Command cmd = new Command("im", "projects");
		
		Response res = mksCmdRunner.execute(cmd);
		if (res != null) {
			WorkItemIterator it = res.getWorkItems();
			while (it.hasNext()) {
				WorkItem wi = it.next();
				String project = wi.getId();
				projects.add(project);
			}
		}
		return projects;
	}
}
