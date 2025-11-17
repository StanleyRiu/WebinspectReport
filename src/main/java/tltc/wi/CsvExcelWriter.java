package tltc.wi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class CsvExcelWriter {
	Configer conf = null;
	String template_file = null;
	FileInputStream fis = null;
	FileOutputStream fos = null;
	XSSFWorkbook xwb = null;
	//XSSFFont xFont = null;
	List<String> eSeverityLevel = Arrays.asList("Critical", "High", "Medium", "Low", "Information", "Best Practice");
//	List<String> cSeverityLevel = Arrays.asList("高風險", "中風險", "低風險", "資訊");
	CsvParser csvParser = null;
	
	public CsvExcelWriter(CsvParser csvParser) {
		this.csvParser = csvParser;
	}
	
	public String write() {
		conf = new Configer();
		template_file = conf.getProperty("XlsxTemplateFile");	//"[源碼][nn]源碼測試總表_ProjectName_YYYYMMDD.xlsx";
		String round = conf.getProperty("Round");
		String version = conf.getProperty("Version");
		String passMsg = "檢測通過";
		String unPassMsg = "檢測尚未通過";
		String noMediumVulnerabilityMsg = "無Medium(含)以上之弱點";
/*
		String InspectPassLevel = conf.getProperty("InspectPassLevel");

		try {
			if (! cSeverityLevel.contains(InspectPassLevel)) {
				throw new Exception("expect InspectPassLevel matches one of \"高風險\", \"中風險\", \"低風險\", \"資訊\"");
			}
		} catch (Exception e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		}
*/		
		try {
			this.fis = new FileInputStream(new File(template_file));
			this.xwb = new XSSFWorkbook(fis);
			//this.xFont = xwb.createFont();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		String projectName = null;
		
	    String s = "://(.+)/?";
	    Pattern pattern = Pattern.compile(s);
	    Matcher matcher = pattern.matcher(csvParser.getStartURL());
	    if (matcher.find()) {
	    	projectName = matcher.group(1).replace("/", "_").replaceFirst(":.*", "");
	    	if (projectName.lastIndexOf("_") == projectName.length() - 1)
	    		projectName = projectName.substring(0, projectName.length() - 1);
	    }

		DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
		String outputDir = new File(csvParser.getCsvFile()).getParent();
		String baseName = new File(outputDir).getName();
		String outputFile = outputDir + File.separator + baseName + "_網站檢測總表_" + formatDate("yyyyMMdd") + ".xlsx";

		replaceTag("Round", round);
		replaceTag("StartURL", csvParser.getStartURL());
		replaceTag("Inspector", conf.getProperty("Inspector"));
		replaceTag("YYYY/MM/DD", formatDate("yyyy/MM/dd"));
		replaceTag("WebinspectVersion", "Webinspect " + conf.getProperty("ScanVersion"));
		replaceTag("JiraId", conf.getProperty("JiraId"));
		replaceTag("SystemVersion", (version.isEmpty() ? formatDate("yyyyMMdd")+"001" : version));
//		replaceTag("Language", xmlDomParser.getLanguage());

		//填寫弱點或通過
		ArrayList<Vulnerability> arrayList = csvParser.getArrayList();
		String startGrid = conf.getProperty("VulnerabilityStartGrid");
		int r = startGrid.codePointAt(1)-'1';
		int c = startGrid.codePointAt(0)-'A';
		
		if (arrayList.size() != 0) {
			//get the events amount in each vulnerabilty
			Iterator<Vulnerability> it = arrayList.iterator();
			while (it.hasNext()) {
				Vulnerability vul = it.next();
//				if (cSeverityLevel.indexOf(vul.getSeverity()) < cSeverityLevel.indexOf(InspectPassLevel))
				vul.setSeverity(eSeverityLevel.get(eSeverityLevel.indexOf(vul.getSeverity())));
				insertVulnerabilityRow(r++, c, vul);
			}
			replaceTag("InspectState", unPassMsg);
		} else {
			//this.xFont.setColor(IndexedColors.BLUE.getIndex());
			this.markNoHighVulnerability(r, c, noMediumVulnerabilityMsg);
			replaceTag("InspectState", passMsg, IndexedColors.BLUE.getIndex());
		}

		this.writeExcelFile(outputFile);
		System.out.println("writing "+outputFile);
		return outputFile;
	}

	private String formatDate(String dateFormat) {
		String format_string = null;
		
		DateFormat df = new SimpleDateFormat(dateFormat);
		format_string = df.format(new java.util.Date());
		return format_string;
	}
	
	private void replaceTag(String findString, String replaceString) {
		this.replaceTag(findString, replaceString, (short) -1);
	}
	
	private void replaceTag(String findString, String replaceString, short indexedColor) {
		XSSFSheet sheet = null;
		XSSFRow xrow = null;
		XSSFCell xcell = null;
		XSSFCellStyle xCellStyle = null;
//		XSSFFont xFont = this.xwb.createFont();
		
//		if (colorIdx != null) xFont.setColor(colorIdx.getIndex()); , IndexedColors colorIdx
		
		sheet = this.xwb.getSheetAt(0);

//		for (Row row : sheet) {
//			for (Cell cell : row) {
//				System.out.println(cell);
//			}
//		}
//		System.out.println(sheet.getPhysicalNumberOfRows());
//		System.out.println(sheet.getFirstRowNum()+" "+sheet.getLastRowNum()+" "+sheet.getTopRow());
		//System.out.println(reportType+" "+cell.getColumnIndex()+" "+cell.getRowIndex());

		for (int r=0; r < sheet.getLastRowNum(); r++) {
			xrow = sheet.getRow(r);
/*
			System.out.print(r+"\t");//xrow.getLastCellNum());
			Iterator<Cell> it = xrow.cellIterator();
			while (it.hasNext()) {
				Cell cell = it.next();
				System.out.print(cell.toString());
			}
			System.out.println();
*/

			for (int c=0; c < xrow.getLastCellNum(); c++) {
				xcell = xrow.getCell(c);
				if (xcell != null) {
					if (xcell.toString().contains(findString)) {
						xCellStyle = xcell.getCellStyle();
						if (indexedColor != -1) xCellStyle.setFont(this.newFontWithColor(xCellStyle, indexedColor));
						xcell.setCellStyle(xCellStyle);
						xcell.setCellValue(xcell.toString().replace(findString, replaceString));
					}
				}
				
					//System.out.println(xrow.getRowNum()+" "+xcell+" removed");
//					sheet.removeRow(xrow);
//					if (r < sheet.getLastRowNum()) {
//						int shiftStartRow = r + 1;
//						int shiftEndRow = sheet.getLastRowNum() < shiftStartRow ? shiftStartRow : sheet.getLastRowNum();  
//						sheet.shiftRows(shiftStartRow, shiftEndRow, -1);
//					}
//				} else r++;

			}
			//System.out.println();
		}


/*
		xrow = sheet.getRow(0);
		for (Cell cell : xrow) {
		}
		for (int r = sheet.getLastRowNum(); r > 0; r--) {
			xrow = sheet.getRow(r);
//			for (int c=0; c < xrow.getPhysicalNumberOfCells(); c++) {
				xcell = xrow.getCell();
//				System.out.println(xrow.getRowNum()+" "+xcell);
				if (xcell == null || xcell.toString().isEmpty()) {
					//System.out.println(xrow.getRowNum()+" "+xcell+" removed");
					sheet.removeRow(xrow);
					if (r < sheet.getLastRowNum()) {
						int shiftStartRow = r + 1;
						int shiftEndRow = sheet.getLastRowNum() < shiftStartRow ? shiftStartRow : sheet.getLastRowNum();  
						sheet.shiftRows(shiftStartRow, shiftEndRow, -1);
					}
				}
//			}
		}
*/		
//		System.out.println(sheet.getPhysicalNumberOfRows());
//		System.out.println(sheet.getFirstRowNum()+" "+sheet.getLastRowNum()+" "+sheet.getTopRow());
		
	}

	public void insertVulnerabilityRow(int r, int c, Vulnerability vul) {
		XSSFSheet xSheet = null;
		XSSFRow xRow = null;
		XSSFCell xCell = null;
		XSSFCellStyle xCellStyleL = null, xCellStyleC = null, xCellStyleR = null;
		
		xSheet = this.xwb.getSheetAt(0);
		xRow = xSheet.getRow(r);
		xCell = xRow.getCell(c);

		xCellStyleL = xCell.getCellStyle().copy();
		xCellStyleC = xCell.getCellStyle().copy();
		xCellStyleR = xCell.getCellStyle().copy();
		xCellStyleL.setAlignment(HorizontalAlignment.LEFT);
		xCellStyleC.setAlignment(HorizontalAlignment.CENTER);
		xCellStyleR.setAlignment(HorizontalAlignment.RIGHT);
		
		xSheet.shiftRows(r, xSheet.getLastRowNum(), 1);
		
		xRow = xSheet.createRow(r);
		xCell = xRow.createCell(c++); xCell.setCellStyle(xCellStyleL); xCell.setCellValue(vul.getName());
		xCell = xRow.createCell(c++); xCell.setCellStyle(xCellStyleR); xCell.setCellValue(vul.getEvents());
		xCell = xRow.createCell(c++); xCell.setCellStyle(xCellStyleC); xCell.setCellValue(vul.getSeverity());
		
		for (int i=0; i<3; i++) {
			xCell = xRow.createCell(c++); 
			xCell.setCellStyle(xCellStyleL);
		}
		CellRangeAddress callRangeAddress = new CellRangeAddress(r,r,4,6);//起始行,結束行,起始列,結束列
		xSheet.addMergedRegion(callRangeAddress);
	}
	
	public void markNoHighVulnerability(int r, int c, String msg) {
		XSSFSheet xSheet = null;
		XSSFRow xRow = null;
		XSSFCell xCell = null;
		XSSFCellStyle xCellStyle = null;
		XSSFFont xFont = null;
		
		xSheet = this.xwb.getSheetAt(0);
		xRow = xSheet.getRow(r);
		xCell = xRow.getCell(c);

		xCellStyle = xCell.getCellStyle();
		
		xFont = this.newFontWithColor(xCellStyle, IndexedColors.BLUE.getIndex());
		
		xCellStyle.setFont(xFont);
		xCell.setCellStyle(xCellStyle);
//		CellUtil.setFont(xCell, xFont);
		xCell.setCellValue(msg);
	}
	
	private XSSFFont newFontWithColor(XSSFCellStyle xCellStyle, short indexedColor)
	{
		XSSFFont xFont = this.xwb.createFont();
		
		xFont.setFontHeight(xCellStyle.getFont().getFontHeight());
		xFont.setFontName(xCellStyle.getFont().getFontName());
		xFont.setColor(indexedColor);
		return xFont;
	}
	
	public OutputStream writeExcelFile(String outFile) {
		File file = new File(outFile);
		//file.getParentFile().mkdirs();
		try {
			this.fos = new FileOutputStream(new File(outFile));
			this.xwb.write(this.fos);
		} catch (IOException e) {
			e.printStackTrace();
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

}
