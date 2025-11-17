package tltc.wi;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

public class CsvParser {
	String csvFile = null;
	Configer conf = null;
	String localPropertiesFile = "local.properties";
	String startURL = null;
	String host = null;
	String scanVersion = null;
	ArrayList<Vulnerability> arrayList = new ArrayList<Vulnerability>();
	
	List<String> severityLevels = Arrays.asList("Critical", "High", "Medium", "Low", "Information", "Best Practice");
	
	Map<String, String> issueHashMap = new HashMap<String, String>();
	SortedMap<String, Integer> vulEvents = new TreeMap<String, Integer>();

	public CsvParser(File file) {
		conf = new Configer(file.getParent()+File.separator+this.localPropertiesFile);
		this.setScanVersion(conf.getProperty("ScanVersion"));
		issueHashMap.put("Critical", "Critical Issues");
		issueHashMap.put("High", "High Issues");
		issueHashMap.put("Medium", "Medium Issues");
		issueHashMap.put("Low", "Low Issues");
		
		this.csvFile = file.getPath();
		String InspectPassLevel = conf.getProperty("InspectPassLevel");
		try {
			if (! severityLevels.contains(InspectPassLevel)) {
				throw new Exception("expect InspectPassLevel matches one of High, Medium, Low, Informational");
			}
		} catch (Exception e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		}

		BufferedReader br = null;
		try {
			br = new BufferedReader(new FileReader(file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		int lineCount = 0;
		String line = null;
		String[] lineArr = null;

		try {
			while ((line = br.readLine()) != null) {
				lineCount++;
				lineArr = line.split(",");
				if (lineCount == 2) {					//csv第二行是scan url
					if (lineArr.length > 1)
						this.setStartURL(lineArr[1]);	
					else this.setStartURL("");			//但是API種類的檢測，.csv檔案沒有列出Scan URL，所以設為""
				}
				if (lineCount <= 8) continue;						//csv第八行開始是弱點
				String Severity = lineArr[0];
				if (! issueHashMap.containsKey(Severity)) continue;	//會有非Critical, High, Medium, Low的訊息，要排除
				String Check = lineArr[1];
				String SeverityCheck = Severity+","+Check;
				
				if (vulEvents.containsKey(SeverityCheck)) vulEvents.replace(SeverityCheck, vulEvents.get(SeverityCheck)+1);
				else vulEvents.put(SeverityCheck, 1);
			}
			
			for (Map.Entry<String, Integer> entry : vulEvents.entrySet()) {
				String[] str = entry.getKey().split(",");
				String severity = str[0];
				if (severityLevels.indexOf(severity) < severityLevels.indexOf(InspectPassLevel)) 
					arrayList.add(new Vulnerability(str[1], entry.getValue(), str[0]));
			}
/*
			vulEvents.forEach((k, v) -> {
				String[] str = k.split(",");
				arrayList.add(new Vulnerability(str[1], v, str[0]));
			});
*/			
			arrayList.sort(new SeveritySort());
//			arrayList.forEach(x -> System.out.println(x.getSeverity()+","+x.getName()+","+x.getEvents()));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public String getCsvFile() {
		return csvFile;
	}

	public void setCsvFile(String csvFile) {
		this.csvFile = csvFile;
	}

	public Configer getConf() {
		return conf;
	}

	public void setConf(Configer conf) {
		this.conf = conf;
	}

	public String getScanVersion() {
		return scanVersion;
	}

	public void setScanVersion(String scanVersion) {
		this.scanVersion = scanVersion;
	}

	public String getStartURL() {
		return startURL;
	}

	public void setStartURL(String startURL) {
		this.startURL = startURL;
	}

	public String getHost() {
		return host;
	}

	public void setHost(String host) {
		this.host = host;
	}

	public ArrayList<Vulnerability> getArrayList() {
		return arrayList;
	}
	
	class SeveritySort implements Comparator<Vulnerability> {
		@Override
		public int compare(Vulnerability o1, Vulnerability o2) {
			int result = severityLevels.indexOf(o1.getSeverity()) - severityLevels.indexOf(o2.getSeverity());
			if (result != 0) return result;
			else return o2.getEvents() - o1.getEvents();
		}
	}
}
