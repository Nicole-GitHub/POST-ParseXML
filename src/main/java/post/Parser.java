package post;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

public class Parser {

	public static void main(String[] args) throws Exception {

		String os = System.getProperty("os.name");

		System.out.println("=== NOW TIME: " + new Date());
		System.out.println("=== os.name: " + os);

		// 判斷當前執行的啟動方式是IDE還是jar
		boolean isStartupFromJar = new File(Parser.class.getProtectionDomain().getCodeSource().getLocation().getPath())
				.isFile();
		System.out.println("=== isStartupFromJar: " + isStartupFromJar);
//		String fileName = "";
		String path = System.getProperty("user.dir") + File.separator; // Jar
		if (!isStartupFromJar) {// IDE
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/POST/JavaTools/POST-ParseXML/" // Mac
					: "C:/Users/nicole_tsou/Dropbox/POST/JavaTools/POST-ParseXML/"; // win

//			fileName = "檔案定義檔.xlsx|檔案定義檔2.xlsx";
		}
		path += "ETL_XML/";
		List<Map<String, String>> listMap = parseXML(path);
		write2Excel(path, listMap);

//		System.out.println("====================== tableMapList =========================");
//		for (Map<String, String> mapList : listMap) {
//			System.out.println("====================== mapList =========================");
//			for (Entry<String, String> set2 : mapList.entrySet()) {
//				System.out.println(set2);
//			}
//		}
		System.out.println("Done !");

	}

	/**
	 * 解析XML內容
	 * 
	 * @param folderPath
	 * @return
	 */
	public static List<Map<String, String>> parseXML(String folderPath) {
		List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
		Map<String, String> map = new HashMap<String, String>();

		String jobName = "", jobActiveStatus = "", frequencyID = "";
		try {

			SAXReader reader = new SAXReader();// 建立解析物件

			// 列出XML清單
			String[] listXML = new File(folderPath).list();
			for (String fileName : listXML) {
				if ("排程整理.xlsx".equals(fileName))
					continue;
				Document document = reader.read(new File(folderPath + fileName));// 載入xml文件
				// Root <JobCategory>
				Element xmlRoot = document.getRootElement();// 得到載入xml文件的根標籤
				Iterator<Element> xmlRootIterator = xmlRoot.elementIterator();// 載入跟標籤以下的所有標籤，返回值是一個迭代器物件
				while (xmlRootIterator.hasNext()) {// 開始遍歷，呼叫hasNext()的方法，判斷是否有下一個
					// 第二層 <JobInfo>
					Element level2 = (Element) xmlRootIterator.next();// 獲取根標籤以下所有子節點，
					Iterator<Element> level2Iterator = level2.elementIterator();// 獲取根節點以下的子子節點，
					while (level2Iterator.hasNext()) {// 遍歷所有的子子節點
						boolean save = true;
						map = new HashMap<String, String>();

						// 第三層 <Auto_Job> <Auto_JobStep>
						Element level3 = (Element) level2Iterator.next();// 獲取子子節點 Element
						Iterator<Element> level3Iterator = level3.elementIterator();// 獲取根節點以下的子子節點，
						while (level3Iterator.hasNext()) {// 遍歷所有的子子節點
							// 第四層 <JobID>
							Element level4 = (Element) level3Iterator.next();// 獲取子子節點 Element
							if ("Auto_Job".equals(level3.getName())) { // 取Job本身
								if ("JobName".equals(level4.getName()))
									jobName = level4.getStringValue();
								if ("FrequencyID".equals(level4.getName()))
									frequencyID = level4.getStringValue();
								if ("ActiveStatus".equals(level4.getName()))
									jobActiveStatus = "0".equals(level4.getStringValue()) ? "N" : "Y";
								save = false;
							} else { // 取JobStep
								if ("StepName".equals(level4.getName()))
									map.put("StepName", level4.getStringValue());
								if ("Description".equals(level4.getName()))
									map.put("Description", level4.getStringValue());
								if ("Command".equals(level4.getName()))
									map.put("Command", level4.getStringValue());
								if ("ActiveStatus".equals(level4.getName()))
									map.put("JobStepActiveStatus", "0".equals(level4.getStringValue()) ? "N" : "Y");
								save = true;
							}
						}

						if (save) {
							map.put("JobCategoryName", fileName);
							map.put("JobName", jobName);
							map.put("FrequencyID", frequencyID);
							map.put("JobActiveStatus", jobActiveStatus);
							listMap.add(map);
						}

					} // 獲取標籤名和標籤中的Text值
				}

			}
		} catch (DocumentException e) {
			e.printStackTrace();
		}

		return listMap;
	}

	/**
	 * 將解析完的XNL內容整理成Excel
	 * 
	 * @param folderPath
	 * @param listMap
	 * @throws Exception
	 */
	private static void write2Excel(String folderPath, List<Map<String, String>> listMap) throws Exception {
		Workbook workbook = Tools.getWorkbook(folderPath + "../Sample - 排程整理.xlsx");
		Sheet sheet = workbook.getSheet("工作表1");
		CellStyle cellStyleNormal = Tools.setStyleNormal(workbook);

		int rownum = 1;
		for (Map<String, String> map : listMap) {
			int cellnum = 0;
			Row row = sheet.createRow(rownum++);
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("FrequencyID"));
			cellnum++; // FREQUENCY_NAME
			cellnum++; // FREQUENCY
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("JobCategoryName"));
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("JobName"));
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("StepName"));
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("Description"));
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("Command"));
			cellnum++; // 前置作業
			cellnum++; // 後續作業
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("JobActiveStatus"));
			Tools.setCell(cellStyleNormal, row, cellnum++, map.get("JobStepActiveStatus"));
		}

		// 將整理好的比對結果另寫出Excel檔
		Tools.output(workbook, folderPath, "排程整理.xlsx");

	}
}
