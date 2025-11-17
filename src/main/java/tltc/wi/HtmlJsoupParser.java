package tltc.wi;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class HtmlJsoupParser {
	String file_name = null;
	Configer conf = new Configer();
	String startURL = null;
	String host = null;
	String scanVersion = null;
	ArrayList<Vulnerability> arrayList = new ArrayList<Vulnerability>();
	List<String> severityLevels = Arrays.asList("Critical", "High", "Medium", "Low", "Info", "BP");
	
	Map<String, String> issueHashMap = new HashMap<String, String>();
	Set<String> issueType = new HashSet<String>();
	Set<String> vulnerability = new HashSet<String>();

	private enum ParseState {
	    READY, ISSUE_TYPE, VULNERABILITY, EVENT
	}
	ParseState parseState;
	
	public HtmlJsoupParser() {
		super();
	}

	public HtmlJsoupParser(File file) {
		issueHashMap.put("Critical", "Critical Issues");
		issueHashMap.put("High", "High Issues");
		issueHashMap.put("Medium", "Medium Issues");
		issueHashMap.put("Low", "Low Issues");
		
		this.file_name = file.getPath();
	}
	
	public void parse() {
		String InspectPassLevel = conf.getProperty("InspectPassLevel");
		try {
			if (! severityLevels.contains(InspectPassLevel)) {
				throw new Exception("expect InspectPassLevel matches one of High, Medium, Low, Informational");
			}
		} catch (Exception e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		}

		Document document = null;
		try {
			document = Jsoup.parse(new File(this.file_name), "UTF-8");
		} catch (IOException e) {
			e.printStackTrace();
		}

		Element scan_version = document.select("div[style=page-break-inside:avoid;page-break-after:always;]").get(1).select("span > nobr").get(3);
		this.setScanVersion(scan_version.text());
		if (scan_version.text() != null) {
			conf.setProperty("ScanVersion", scan_version.text());
			conf.saveProperties();
			return;	//get ScanVersion only for html file, using parse csv instead
		}
		
		Element startUrl = document.select("span[style$=background-color:#4F81BD;] + span[style$=background-color:#DBE5F1;] > nobr:containsOwn(http)").get(0);
		this.setStartURL(startUrl.text());

//		Element finding = document.select("span[style$=background-color:#4F81BD;] + span[style$=background-color:#DBE5F1;] > nobr:containsOwn(http)").get(0);
//		System.out.println(scan_version.text());
//		System.exit(0);
//		Elements findings = document.select("div[style=page-break-inside:avoid;page-break-after:always;]");
//		for (Iterator<Element> it = findings.iterator(); it.hasNext(); ) {
//			Element element = it.next();
//			System.out.println(element.text());
//		}
//		System.exit(0);

/*		
		Element host = document.select("table.ax-scan-summary > tbody > tr").get(3).select("td").get(1);
		this.setHost(host.text());
*/
		Elements issues = document.select("span[style$=background-color:#4F81BD;] > a[name] + nobr:containsOwn(Issues)");	//get Critical, High, Medium Issue Types
		for (Iterator<Element> it = issues.iterator(); it.hasNext(); ) {
			Element element = it.next();
			this.issueType.add(element.text());
		}

		Elements vulnerabilities = document.select("span[style$=background-color:#DBE5F1;] > a[name] + nobr:matchesOwn(\\(.*\\d+.*\\))");	//get vulnerabilities
		for (Iterator<Element> it = vulnerabilities.iterator(); it.hasNext(); ) {
			Element element = it.next();
			this.vulnerability.add(element.text());
			System.out.println(element.text());
		}
//		vulnerabilities.forEach(x -> System.out.println(x.text()));

//		Elements issues = document.select("span[style$=font-weight:bold;] > nobr:containsOwn(Page:)");	//get vulnerability counts

		String severity = null;
		String vulerability = null;
		int events = 0;

		Elements issues_vuls_events = document.select("nobr:containsOwn(Issues), span[style$=background-color:#DBE5F1;] > a[name] + nobr:matchesOwn(\\(.*\\d+.*\\)), nobr:containsOwn(Page:)");
		for (Iterator<Element> it = issues_vuls_events.iterator(); it.hasNext(); ) {
			Element element = it.next();

			if (this.issueHashMap.containsValue(element.text())) {
				if (parseState == ParseState.EVENT) arrayList.add(new Vulnerability(vulerability, events, severity));
				parseState = ParseState.ISSUE_TYPE;
				for (Map.Entry<String, String> entry : issueHashMap.entrySet()) {
					if (entry.getValue().equals(element.text())) {
						severity = entry.getKey();
					}
				}
			}

		    String p = "\\(.*\\d+.*\\)";
		    Pattern pattern = Pattern.compile(p);
		    Matcher matcher = pattern.matcher(element.text());
		    if (matcher.find()) {
				if (parseState == ParseState.EVENT && severityLevels.indexOf(severity) < severityLevels.indexOf(InspectPassLevel)) {
					arrayList.add(new Vulnerability(vulerability, events, severity));
				}
		    	parseState = ParseState.VULNERABILITY;
		    	vulerability = element.text();
				events = 0;
		    }

			if (element.text().equalsIgnoreCase("Page:")) {
				parseState = ParseState.EVENT;
				events++;
			}
			if (! it.hasNext()) arrayList.add(new Vulnerability(vulerability, events, severity));
		}

//		arrayList.forEach(x -> System.out.println(x.getName()+", "+x.getEvents()+", "+x.getSeverity()));
/*
		//get ax-alerts-distribution
		Element table = document.select("table.ax-alerts-distribution").first();	//only one
		Elements rows = table.select("tr");
		for (Iterator<Element> it = rows.iterator(); it.hasNext();) {
	        Element element = it.next();
	        // do something
//        	System.out.println(element.text());
	        String[] alert = element.text().split(" +");
        	if (eSeverityLevel.contains(alert[0]) && eSeverityLevel.indexOf(alert[0]) < 2) { 
        		alertsHashMap.put(alert[0], alert[1]);
        	}
	    }
*/		
//		alertsHashMap.forEach((k, v) -> System.out.println(k + ":" + v));
//		arrayList.forEach(x -> System.out.println(x.getName()+" "+x.getEvents()+" "+x.getSeverity()));
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
}


