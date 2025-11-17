package tltc.wi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;

public class WordWriter {
	Configer conf = null;
	String template_file = null;
	FileInputStream fis = null;
	FileOutputStream fos = null;
	XSSFWorkbook xwb = null;
	XWPFDocument docx = null;
	//XSSFFont xFont = null;
	CsvParser csvParser = null;
	List<String> eSeverityLevel = Arrays.asList("High", "Medium", "Low", "Informational");
	String passMsg = "檢測通過";
	String unPassMsg = "檢測尚未通過";
	String[] subProjectName = null;
	String InspectPassLevel = null; 
	String version = null;

	public WordWriter(CsvParser csvParser) {
		this.csvParser = csvParser;
	}
	
	public String write() {
		conf = csvParser.getConf();
		InspectPassLevel = conf.getProperty("InspectPassLevel");
		version = conf.getProperty("Version");
		template_file = conf.getProperty("DocxTemplateFile");	//"[網頁][nn]網頁測試總表_Host_YYYYMMDD.xlsx";
		String module = "A";
		String round = conf.getProperty("Round");
		String noMediumVulnerabilityMsg = "無高於通過標準以上之弱點";

		try {
			this.fis = new FileInputStream(new File(template_file));
//	toURI()的寫法在eclipse IDE中沒問題，但是在command line中會出問題
//			this.fis = new FileInputStream(new File(this.getClass().getClassLoader().getResource("."+File.separator+template_file).toURI()));
//			this.fis = (FileInputStream) this.getClass().getClassLoader().getResourceAsStream("."+File.separator+template_file);
			this.docx = new XWPFDocument(fis);
			//this.xFont = xwb.createFont();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
		String outputDir = new File(csvParser.getCsvFile()).getParent();
		String baseName = new File(outputDir).getName();
		String outputFile = outputDir + File.separator + baseName + "_網站檢測總表_" + formatDate("yyyyMMdd") + ".docx";
		subProjectName = baseName.split("(\\[.+\\])+");

		//填寫弱點或通過
		ArrayList<Vulnerability> arrayList = csvParser.getArrayList();	//所有弱點都先解析放入ArrayList中
		
		XWPFTable table;
		//處理第一個表格
		table = this.docx.getTableArray(0);	//最新檢測結果摘要
		printSummary(table, arrayList);

		//處理第二個表格
		table = this.docx.getTableArray(1);	//弱點風險說明單

		if (arrayList.size() > 0) {
			printVulnerabilityHeader(table, arrayList);
			fillExemption(table, arrayList);
		} else {
			//remove page two
			XWPFParagraph paragraph = this.docx.getParagraphs().stream()
					.filter(p -> p.getParagraphText().contains("弱點風險說明單"))
					.findFirst().orElse(null);
			if (paragraph != null) {
				this.docx.removeBodyElement(this.docx.getPosOfParagraph(paragraph));
				this.docx.removeBodyElement(this.docx.getPosOfTable(table));
			}
		}
		
		//處理第三個表格以後
		//附件、檢測歷程
//		table = this.docx.getTableArray(2);	//附件、檢測歷程
//
//	    XWPFParagraph para = this.docx.getParagraphs().stream()
//	            .filter(p -> p.getParagraphText().contains("檢測歷程"))
//	            .findFirst().orElse(null);
//	    if (para != null) {
//			this.docx.removeBodyElement(this.docx.getPosOfParagraph(para));
//			this.docx.removeBodyElement(this.docx.getPosOfTable(table));
//	    }
	    //刪除paragraph、table後還有三個空白段落
	    this.docx.removeBodyElement(this.docx.getPosOfParagraph(this.docx.getLastParagraph()));
//	    this.docx.removeBodyElement(this.docx.getPosOfParagraph(this.docx.getLastParagraph()));
//	    this.docx.removeBodyElement(this.docx.getPosOfParagraph(this.docx.getLastParagraph()));
		
		this.writeWordFile(outputFile);
		System.out.println("writing "+outputFile);
		return outputFile;
	}

	private void printSummary(XWPFTable table, ArrayList<Vulnerability> arrayList) {
		for (XWPFTableRow tableRow : table.getRows()) {
			for (XWPFTableCell tableCell : tableRow.getTableCells()) {
				for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
					replaceTextInParagraph(paragraph, "ProjectName", subProjectName[1]);//this.xmlDomParser.getProjectName());
					replaceTextInParagraph(paragraph, "YYYYMMDD", formatDate("yyyy/MM/dd"), "0000FF");
					replaceTextInParagraph(paragraph, "Inspector", conf.getProperty("Inspector"));
					replaceTextInParagraph(paragraph, "UnpassLevel", eSeverityLevel.get(eSeverityLevel.indexOf(InspectPassLevel) - 1));
					replaceTextInParagraph(paragraph, "PassLevel", conf.getProperty("InspectPassLevel"));
					if (arrayList.size() != 0) {
						replaceTextInParagraph(paragraph, "InspectResult", unPassMsg, "FF0000");
					} else {
						replaceTextInParagraph(paragraph, "InspectResult", passMsg, "0000FF");
					}
					replaceTextInParagraph(paragraph, "StartURL", csvParser.getStartURL());
					replaceTextInParagraph(paragraph, "WebinspectVersion", "Webinspect " + csvParser.getScanVersion());
					replaceTextInParagraph(paragraph, "SystemVersion", (version.isEmpty() ? formatDate("yyyyMMdd") + "001" : version));
					replaceTextInParagraph(paragraph, "TaskId", conf.getProperty("JiraId"));
				}
			}
		}
		
		if (arrayList.size() > 0) {
			addVulnerability(table, arrayList);
		}
	}

	private void printVulnerabilityHeader(XWPFTable table, ArrayList<Vulnerability> arrayList) {
		for (XWPFTableRow tableRow : table.getRows()) {
			for (XWPFTableCell tableCell : tableRow.getTableCells()) {
				for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
					replaceTextInParagraph(paragraph, "CustomerName", this.conf.getProperty("customerName"));
					replaceTextInParagraph(paragraph, "ProjectName", this.conf.getProperty("projectName"));
					replaceTextInParagraph(paragraph, "TaskId", conf.getProperty("JiraId"));
				}
			}
		}
//		fillExemption(table, arrayList);
	}
	
	private void commitTableRows(XWPFTable table) {
		int rowNr = 0;
		for (XWPFTableRow tableRow : table.getRows()) {
			table.getCTTbl().setTrArray(rowNr++, tableRow.getCtRow());
		}
	}
	 
	private void fillExemption(XWPFTable table, ArrayList<Vulnerability> arrayList) {
		int pos = 0;
		boolean bFound = false;
		List<XWPFTableRow> tableRows = table.getRows();
		
		for (XWPFTableRow tableRow : tableRows) {
			for (XWPFTableCell tableCell : tableRow.getTableCells()) {
				if (tableCell.getText().equals("風險說明列表")) {
					pos = tableRows.indexOf(tableRow);
					bFound = true;
					break;
				}
			}
			if (bFound) break;
		}
		
	    CTRPr ctRPr = CTRPr.Factory.newInstance();
        CTFonts ctFonts = ctRPr.addNewRFonts();
        ctFonts.setAscii("Arial");
        ctFonts.setCs("Arial");
        ctFonts.setEastAsia("標楷體");
        ctFonts.setHAnsi("Arial");
        ctRPr.addNewSz().setVal(new BigInteger("22"));
        ctRPr.addNewSzCs().setVal(new BigInteger("22"));

		int step = 3;
		
		for (int i = 0; i < arrayList.size(); i++) {
			Vulnerability vul = arrayList.get(i);
			if (arrayList.size() - i > 1)
				replicateRows(table, pos + i*step + 1, pos + i*step + 4);
			XWPFTableRow row = table.getRow(pos + i*step + 1);
			XWPFTableCell cell;
			XWPFParagraph para;
			XWPFRun run;
			
			cell = row.getCell(0);	//弱點序號
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(String.valueOf(i+1), 0);
			ctRPr.getSzArray(0).setVal(new BigInteger("22"));
	        ctRPr.getSzCsArray(0).setVal(new BigInteger("22"));
			run.getCTR().setRPr(ctRPr);

			cell = row.getCell(2);	//弱點名稱
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.LEFT);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(vul.getName(), 0);
			ctRPr.getSzArray(0).setVal(new BigInteger("20"));
	        ctRPr.getSzCsArray(0).setVal(new BigInteger("20"));
			run.getCTR().setRPr(ctRPr);
			
			row = table.getRow(pos + i*step + 2);
			cell = row.getCell(2);	//數量
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(String.valueOf(vul.getEvents()), 0);
			run.getCTR().setRPr(ctRPr);
			
			cell = row.getCell(4);	//弱點等級
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(vul.getSeverity(), 0);
			run.getCTR().setRPr(ctRPr);

			commitTableRows(table);
		}
	}
	
	private void replicateRows(XWPFTable table, int beginRowId, int endRowId) {
		CTRow ctRow = null;
		
		for (int i = 0; i < endRowId - beginRowId; i++ ) {
			try {
				ctRow = CTRow.Factory.parse(table.getRow(beginRowId + i).getCtRow().newInputStream());
			} catch (XmlException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (NullPointerException e) {
				System.out.println(e.getMessage());
			}
			
		    CTRPr ctRPr = CTRPr.Factory.newInstance();
	        CTFonts ctFonts = ctRPr.addNewRFonts();
	        ctFonts.setAscii("Arial");
	        ctFonts.setCs("Arial");
	        ctFonts.setEastAsia("標楷體");
	        ctFonts.setHAnsi("Arial");
	        ctRPr.addNewSz().setVal(new BigInteger("22"));
	        ctRPr.addNewSzCs().setVal(new BigInteger("22"));

			XWPFTableRow newRow;
			XWPFTableCell cell;
			XWPFParagraph para;
			XWPFRun run;
			
			newRow = new XWPFTableRow(ctRow, table);
			for (XWPFTableCell tableCell : newRow.getTableCells()) {
//				while (tableCell.getParagraphs().size() > 0) tableCell.removeParagraph(0);
			}
//			para = cell.addParagraph();
			table.addRow(newRow, endRowId + i);
		}
	}
	
	private void addVulnerability(XWPFTable table, ArrayList<Vulnerability> arrayList) {
		int pos = 0;
		boolean bFound = false;
		List<XWPFTableRow> tableRows = table.getRows();
		
		for (XWPFTableRow tableRow : tableRows) {
			for (XWPFTableCell tableCell : tableRow.getTableCells()) {
				if (tableCell.getText().equals("弱點名稱")) {
					pos = tableRows.indexOf(tableRow);
					bFound = true;
					break;
				}
			}
			if (bFound) break;
		}

		CTRow ctRow = null;
		try {
			ctRow = CTRow.Factory.parse(table.getRow(pos+1).getCtRow().newInputStream());
		} catch (XmlException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		XWPFTableRow newRow;// = new XWPFTableRow(ctRow, table);

	    CTRPr ctRPr = CTRPr.Factory.newInstance();
        CTFonts ctFonts = ctRPr.addNewRFonts();
        ctFonts.setAscii("Arial");
        ctFonts.setCs("Arial");
        ctFonts.setEastAsia("標楷體");
        ctFonts.setHAnsi("Arial");
        ctRPr.addNewSz().setVal(new BigInteger("22"));
        ctRPr.addNewSzCs().setVal(new BigInteger("22"));

		int i = 1;
		for (Vulnerability vul : arrayList) {
			newRow = new XWPFTableRow(ctRow, table);

			XWPFTableCell cell;
			XWPFParagraph para;
			XWPFRun run;
			cell = newRow.getCell(0);
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(vul.getName(), 0);
			ctRPr.getSzArray(0).setVal(new BigInteger("20"));
	        ctRPr.getSzCsArray(0).setVal(new BigInteger("20"));
			run.getCTR().setRPr(ctRPr);
			
			cell = newRow.getCell(1);
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.RIGHT);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(String.valueOf(vul.getEvents()), 0);
			run.getCTR().setRPr(ctRPr);
			
			cell = newRow.getCell(2);
			while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
			para = cell.addParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
//			para = cell.getParagraphArray(0);
//			while (para.getRuns().size() > 0) para.removeRun(0);
			run = para.createRun();
			run.setText(vul.getSeverity(), 0);
			run.getCTR().setRPr(ctRPr);
			
			table.addRow(newRow, ++pos);
		}
	}
	
 	private void replaceTextInParagraph(XWPFParagraph paragraph, String originalText, String replaceText) {
 		replaceTextInParagraph(paragraph, originalText, replaceText, "000000");
	}
	
	private void replaceTextInParagraph(XWPFParagraph paragraph, String originalText, String replaceText, String color) {
	    Double fontSize = 0.0;
	    String fontFamily;
		String paragraphText = paragraph.getParagraphText();
	    TextSegment textSegment;
	    int position = 0, begin, end;

	    CTRPr ctRPr = CTRPr.Factory.newInstance();
        CTFonts ctFonts = ctRPr.addNewRFonts();
        ctFonts.setAscii("Arial");
        ctFonts.setCs("Arial");
        ctFonts.setEastAsia("標楷體");
        ctFonts.setHAnsi("Arial");
        ctRPr.addNewSz().setVal(new BigInteger("24"));
        ctRPr.addNewSzCs().setVal(new BigInteger("24"));

	    if (null != (textSegment = paragraph.searchText(originalText, new PositionInParagraph(position, 0, 0)))) {
	    	
	    	begin = textSegment.getBeginRun();
	    	end = textSegment.getEndRun();
//	    	System.out.println(begin+" "+end);
	    	CTParaRPr ctParaRPr = paragraph.getCTP().getPPr().getRPr();
	    	CTPPr ctPPr = paragraph.getCTP().getPPr();
	    	paragraph.getCTP().setPPr(ctPPr);
	    	XWPFRun run = paragraph.getRuns().get(begin);
	    	ctRPr = run.getCTR().getRPr();

        	//add new run with updated text
            XWPFRun newRun = paragraph.insertNewRun(begin);	//createRun();
            newRun.getCTR().setRPr(ctRPr);
            newRun.setText(replaceText);//newText);
            newRun.setColor(color);
//            newRun.setFontFamily("標楷體");
//            newRun.setFontSize(12);
//            paragraph.addRun(run);

	    	int cnt = end - begin + 1;
	    	for (int i = 0; i < cnt; i++) paragraph.removeRun(begin+1);

	    }	
	}
	
	private String formatDate(String dateFormat) {
		String format_string = null;
		
		DateFormat df = new SimpleDateFormat(dateFormat);
		format_string = df.format(new java.util.Date());
		return format_string;
	}
	
	public OutputStream writeWordFile(String outFile) {
		File file = new File(outFile);
		//file.getParentFile().mkdirs();
		try {
			this.fos = new FileOutputStream(new File(outFile));
			this.docx.write(this.fos);
		} catch (IOException e) {
			System.err.println(e.getMessage());
		}
		finally {
			if (this.fos != null)
				try {
					this.fos.flush();
					this.fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			if (this.xwb != null) {
				try {
					this.xwb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return this.fos;
	}

	private void replaceTextFor(XWPFTable table, HashMap<?, ?> map) {
		table.getRows().forEach(r -> r.getTableCells().forEach(c -> 
		c.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains((String) findText)) {
                    run.setText(text.replace((String) findText, (String) replaceText), 0);
                }
            });
        }))));
    }

	private void replaceTextFor(XWPFParagraph paragraph, HashMap<?, ?> map) {
		paragraph.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains((String) findText)) {
                    run.setText(text.replace((String) findText, (String) replaceText), 0);
                }
            });
        });
    }
	
	private void replaceTextFor(XWPFDocument doc, HashMap<?, ?> map) {
        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains((String) findText)) {
                    run.setText(text.replace((String) findText, (String) replaceText), 0);
                }
            });
        }));
    }
}
