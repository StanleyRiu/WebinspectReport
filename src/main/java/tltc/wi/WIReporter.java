package tltc.wi;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class WIReporter {
	File[] files = null;

	public WIReporter(String[] args) {
		if (args.length != 0) {
			List<File> list = new ArrayList<File>();
			for (String s : args) {
				if (s.toLowerCase().endsWith(".csv")) {
					list.add(new File(s));
				}
			}
			if (! list.isEmpty()) files = list.toArray(new File[list.size()]);
			else {
				System.out.println("no .csv files found, exit");
				System.exit(0);
			}
		}
		if (files == null) {
			files = this.listFiles(".");
		}
	}

	public static void main(String[] args) {
		WIReporter wiReporter = new WIReporter(args);
		wiReporter.generateReports();
	}	
	
	public List<String> generateReports() {
		List<String> list = new ArrayList<String>();
		for (File file : this.files) {
			System.out.println("parsing "+file.getName()+"...");
//			HtmlJsoupParser htmlDomParser = new HtmlJsoupParser(file);
//			htmlDomParser.parse();
//			ExcelWriter excelWriter = new ExcelWriter(htmlDomParser);
			CsvParser csvParser = new CsvParser(file);
			String excelFileName = new CsvExcelWriter(csvParser).write();
			list.add(excelFileName);
			String wordFileName = new WordWriter(csvParser).write();
			list.add(wordFileName);
		}
		return list;
	}

	public File[] listFiles(String path)	{
		File file = new File(path);
		File[] files = null;

		Filter filter = new Filter();
		files = file.listFiles(filter);
		return files;
	}
	
	private class Filter implements FilenameFilter {
		@Override
		public boolean accept(File dir, String name) {
            if(name.lastIndexOf('.') > 0) {
               // get last index for '.' char
               int lastIndex = name.lastIndexOf('.');
               
               // get extension
               String str = name.substring(lastIndex);
               
               // match path name extension
               if(str.equalsIgnoreCase(".html")) {
                  return true;
               }
            }
            return false;
		}
		
	}

//	private void zip() {
//		String[] files = new String[] {};
//		files = Stream.of(this.officeFileNames, this.watermarkedPDFs).flatMap(Stream::of).collect(Collectors.toList()).toArray(files);
//
//		File f = new File(this.ReportName).getParentFile();
//		String[] args = new String[] {"-p", conf.getProperty("EncryptPassword"), "-f", f.toString()+File.separator+f.getName()+".zip" };
//		List<String> list = new ArrayList<>(Arrays.asList(args));
//		List<String> lst = Arrays.asList(files); 
//		list.addAll(lst);
//		
////		Arrays.stream(args).forEach(System.out::println);
//		Zipper zipper = new Zipper();
//		args = list.toArray(args);
//		zipper.cli(args);
//		System.out.println("finished zipping");
//	}

}
