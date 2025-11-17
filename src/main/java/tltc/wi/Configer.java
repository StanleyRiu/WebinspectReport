package tltc.wi;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.Properties;
import java.util.Scanner;

public class Configer extends Properties {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	String globalPropertiesFile = "wiconfig.properties";
	String localPropertiesFile = null;

	public Configer() {
		loadProperties(this.globalPropertiesFile);
	}

	public Configer(String propertiesFile) {
		loadProperties(this.globalPropertiesFile);

		this.localPropertiesFile = propertiesFile;

		//later loaded key will replace the same key loaded before
		File localPropFile = new File(this.localPropertiesFile);
		if (localPropFile.exists() && ! localPropFile.isDirectory()) {
			loadProperties(this.localPropertiesFile);
		}
	}

	void loadProperties(String propertiesFile) {
		try {
			//if this.getClass().getClassLoader().getResourceAsStream() return null, that is CLASSPATH lacks of path of resource
////			InputStream pis = this.getClass().getClassLoader().getResourceAsStream(propertiesFile);
//			InputStream pis = Thread.currentThread().getContextClassLoader().getResourceAsStream(propertiesFile);
			File file = new File(propertiesFile);
			Scanner scanner = new Scanner(file, "UTF-8");
//			Scanner scanner = new Scanner(new File(propertiesFile), "UTF-8");
//			scanner.forEachRemaining(System.out::println);
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			while (scanner.hasNext()) {
				String line = scanner.nextLine();
				baos.write(line.replace(File.separator, File.separator.repeat(2)).getBytes());
				baos.write(System.lineSeparator().getBytes());
			}
			InputStream is = new ByteArrayInputStream(baos.toByteArray());
			this.load(new InputStreamReader(is));//, "UTF-8"));//don't assign charset again, since Scanner has assign charset
//			this.load(new BufferedReader(new FileReader(propertiesFile, Charset.forName("UTF-8"))));
		} catch (IOException e) {
			e.printStackTrace();
		}
//		this.list(System.out);
//		System.exit(0);
	}
	
	public String getProperty(String key) {
		//OSSFile ossfile = new OSSFile();
		//ossfile.listFiles();
	
	//	InputStream is = null;
	//	is = this.getClass().getClassLoader().getResourceAsStream(propertiesFile);
	//	if (is != null) {
	//		prop.load(is);
	//	} else {
	//		throw new FileNotFoundException("property file '" + propertiesFile + "' not found in the classpath");
	//	}
		
		String value = null;
		value = super.getProperty(key); 
		if (value == null) {
			value = "";
			try {
				throw new Exception("key: "+key+" or value not found");
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return value;
	}

	public void saveProperties() {
		//OutputStream os = new FileOutputStream(propertiesFile);
		BufferedWriter bw;
		
		try {
			bw = new BufferedWriter(new FileWriter(this.globalPropertiesFile, Charset.forName("UTF-8")));
			this.store(bw, null);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
