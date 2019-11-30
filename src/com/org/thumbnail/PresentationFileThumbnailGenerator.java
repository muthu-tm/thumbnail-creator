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
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.jodconverter.LocalConverter;
import org.jodconverter.LocalConverter.Builder;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;

public class PresentationFileThumbnailGenerator {
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
		if (extension.equalsIgnoreCase("ppt")) {
			count = getPPTPageCount(inputFile);
		} else if (extension.equalsIgnoreCase("pptx")) {
			count = getPPTXPageCount(inputFile);
		}
		File tempFile = new File(TEMP + "/" + inputFile.getName());
		try {
			int page = 1;
			while (count > 0) {
				if (page == 1) {
					FileUtils.copyFile(inputFile, tempFile);
				}
				count--;
				File input = tempFile;
				Builder builder = LocalConverter.builder();
				File outputFile = new File(thumbnailPath + "page_" + page + imagePrefix);
				builder.build().convert(input).to(outputFile).execute();
				if (extension.equalsIgnoreCase("ppt")) {
					removePPTPage(tempFile);
				} else if (extension.equalsIgnoreCase("pptx")) {
					removePPTXPage(tempFile);
				}
				page++;
			}
		} catch (OfficeException | IOException e) {
			System.out.println("OfficeException/IOException caught while processing page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while processing page:" + e.getMessage());
		} finally {
			tempFile.deleteOnExit();
		}
	}

	private int getPPTPageCount(File file) {
		int pageCount = 0;
		try (HSLFSlideShow document = new HSLFSlideShow(new FileInputStream(file))) {
			pageCount = document.getSlides().size();
		} catch (IOException e) {
			System.out.println("IOException caught while getting PPT page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while getting PPT page:" + e.getMessage());
		} 
		return pageCount;
	}

	private int getPPTXPageCount(File file) {
		int pageCount = 0;
		try (XMLSlideShow xslideShow = new XMLSlideShow(OPCPackage.open(file))) {
			pageCount = xslideShow.getSlides().size();
		} catch (OpenXML4JException e) {
			System.out.println("IOException caught while getting PPTX page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while getting PPTX page:" + e.getMessage());
		}
		return pageCount;
	}

	private void removePPTPage(File source) throws IOException {
		try (HSLFSlideShow document = new HSLFSlideShow(new FileInputStream(source))) {
			document.removeSlide(0);
			source.delete();
			FileOutputStream out = new FileOutputStream(new File(TEMP + "/" + source.getName()));
			document.write(out);
		} catch (IOException e) {
			System.out.println("IOException caught while removing PPT page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while removing PPT page:" + e.getMessage());
		}
	}

	private void removePPTXPage(File source) throws IOException {
		try (XMLSlideShow xslideShow = new XMLSlideShow(OPCPackage.open(source))) {
			xslideShow.removeSlide(0);
			source.delete();
			FileOutputStream out = new FileOutputStream(new File(TEMP + "/" + source.getName()));
			xslideShow.write(out);
		} catch (OpenXML4JException e) {
			System.out.println("OpenXML4JException caught while removing PPTX page:" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught while removing PPTX page:" + e.getMessage());
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

		PresentationFileThumbnailGenerator pptThumbnail = new PresentationFileThumbnailGenerator();
		final LocalOfficeManager officeManager = LocalOfficeManager.builder().officeHome(getOfficeHome()).install()
				.build();
		try {
			officeManager.start();
			pptThumbnail.generate(pptThumbnail.loadConfig(filePath.toString()));
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
