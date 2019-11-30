package com.org.thumbnail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jodconverter.LocalConverter;
import org.jodconverter.LocalConverter.Builder;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;

public class ExcelThumbnailGenerator {
	
	String TEMP = "";
	
	public void generate(Properties prop) {
		String filePath = prop.getProperty("SorceFile");
		String thumbnailPath = prop.getProperty("OutputFolder");
		String imagePrefix = prop.getProperty("Format");
		TEMP = prop.getProperty("tempFolder");
		
		File newFolder = new File(thumbnailPath);
		boolean created = newFolder.mkdirs();
		File folder = new File(TEMP);
		created = folder.mkdirs();		
		
		File inputFile = new File(filePath);
		String[] filePathArr = filePath.split("\\.");
		String extension = filePathArr[filePathArr.length - 1];
		int count = 0;
		if (extension.equalsIgnoreCase("xls")) {
			count = getXLSPageCount(inputFile);
		} else if (extension.equalsIgnoreCase("xlsx")) {
			count = getXLSXPageCount(inputFile);
		}
		File tempFile = new File(TEMP + "/" + inputFile.getName());
		try {
			int page = 1;
			while (count > 0) {
				if (page == 1) {
					FileUtils.copyFile(inputFile, tempFile);
				}

				File input = tempFile;
				Builder builder = LocalConverter.builder();
				File outputFile = new File(thumbnailPath + "page_" + page + imagePrefix);
				builder.build().convert(input).to(outputFile).execute();
				if (extension.equalsIgnoreCase("xls")) {
					removeXLSPage(tempFile);
				} else if (extension.equalsIgnoreCase("xlsx")) {
					removeXLSXPage(tempFile);
				}

				page++;
				count--;
			}
		} catch (OfficeException | IOException e) {
			System.out.println("OfficeException/IOException caught while processing page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while processing page:" + e.getMessage());
		} finally {
			System.out.println("Delete Temp File");
			tempFile.deleteOnExit();
		}
	}

	private int getXLSXPageCount(File file) {
		int pageCount = 0;
		try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
			pageCount = workbook.getNumberOfSheets();
		} catch (IOException e) {
			System.out.println("IOException caught while loading XLSX page:::" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while getting XLSX page - " + e.getMessage());
		}
		return pageCount;
	}

	private int getXLSPageCount(File file) {
		int pageCount = 0;
		try (HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file))) {
			pageCount = workbook.getNumberOfSheets();
		} catch (OldExcelFormatException e) {
			System.out.println("OldExcelFormatException caught while loading XLS page:::" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while getting XLS page - " + e.getMessage());
		}
		return pageCount;
	}

	private void removeXLSXPage(File source) throws IOException {
		try(XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(source))) {
			workbook.removeSheetAt(0);
			source.delete();
			FileOutputStream out = new FileOutputStream(new File(TEMP + "/" + source.getName()));
			workbook.write(out);
		} catch (IOException e) {
			System.out.println("I/O Exception caught while removing XLSX page - "+ e.toString());
		} catch (Exception e) {
			System.out.println("Exception caught while removing XLSX page - " + e.getMessage());
		}
	}

	private void removeXLSPage(File source) throws IOException {
		try(HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(source))) {
			workbook.removeSheetAt(0);
			source.delete();
			FileOutputStream out = new FileOutputStream(new File(TEMP + "/" + source.getName()));
			workbook.write(out);
		} catch (IOException e) {
			System.out.println("I/O Exception caught while removing XLS page - "+ e.toString());
		} catch (Exception e) {
			System.out.println("Exception caught while removing XLS page - " + e.getMessage());
		}
	}

	private Properties loadConfig(String filePath) throws IOException {
		Properties prop = new Properties();
		FileReader reader = new FileReader(filePath);
		prop.load(reader);
		return prop;
	}

	public static void main(String[] args) throws IOException {
		Path currentPath = Paths.get(System.getProperty("user.dir"));
		Path filePath = Paths.get(currentPath.toString(), "resources", "config.properties");

		ExcelThumbnailGenerator excelThumbnail = new ExcelThumbnailGenerator();
		final LocalOfficeManager officeManager = LocalOfficeManager.builder().officeHome(getOfficeHome()).install()
				.build();
		try {
			officeManager.start();
			excelThumbnail.generate(excelThumbnail.loadConfig(filePath.toString()));
		} catch (OfficeException e) {
			e.printStackTrace();
		} finally {
			System.out.println("Completed Processing; stopping officeManager!");
			OfficeUtils.stopQuietly(officeManager);
		}
	}
	
	/**
	 * Provides the libreoffice home path, if it's wrong JODconverter will throw
	 * error "officeHome not set and could not be auto-detected"; for Mac OSX ->
	 * "/Applications/LibreOffice.app/Contents" for linux -> "/opt/libreoffice6.3"
	 * 
	 * @return {@link String} officeHome - libreoffice home path
	 */
	private static String getOfficeHome() {
		String os = System.getProperty("os.name").toLowerCase();
		String officeHome = "";

		if (os.indexOf("nix") >= 0 || os.indexOf("nux") >= 0 || os.indexOf("aix") > 0) {
			// Unix or Linux sstem office home path
			officeHome = "/opt/libreoffice";
		} else if (os.indexOf("mac") >= 0) {
			// Mac os office home path
			officeHome = "/Applications/LibreOffice.app/Contents";
		} else if (os.indexOf("win") >= 0) {
			// Windows system office home path
			officeHome = "C:\\Program Files (x86)\\LibreOffice 6.3";
		} else if (os.indexOf("sunos") >= 0) {
			// Solaris system office home path
		}

		return officeHome;
	}

}
