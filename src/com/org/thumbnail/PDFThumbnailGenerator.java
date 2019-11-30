package com.org.thumbnail;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;

public class PDFThumbnailGenerator {
	
	public void generate(Properties prop) {
		String filePath = prop.getProperty("SorceFile");
		String thumbnailPath = prop.getProperty("OutputFolder");
		String imagePrefix = prop.getProperty("Format");

		File newFolder = new File(thumbnailPath);
		boolean created = newFolder.mkdirs();

		File inputFile = new File(filePath);
		PDDocument document = null;
		try {
			document = PDDocument.load(inputFile);
			PDFRenderer renderer = new PDFRenderer(document);
			int page = 0;
			
			for (int index = 0; index < document.getNumberOfPages(); index++) {
				++page;
				File imageFile = new File(thumbnailPath + "page_" + page + imagePrefix);
				BufferedImage image = renderer.renderImageWithDPI(index, 110);
				
				// Create thumbnail for current page
				ImageIO.write(image, "png", imageFile);
			}
			document.close();
		} catch (IOException e) {
			System.out.println("IOException caught while processing PDF file::" + e.getMessage());
		} catch (Exception e) {
			System.out.println("Exception caught from here:::" + e.getMessage());
		} finally {
			try {
				if (document != null)
					document.close();
			} catch (IOException e) {
				System.out.println("IOException caught while closing resource:" + e.getMessage());
			}
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

		PDFThumbnailGenerator pdfThumbnail = new PDFThumbnailGenerator();
		final LocalOfficeManager officeManager = LocalOfficeManager.builder().officeHome(getOfficeHome()).install()
				.build();
		try {
			officeManager.start();
			pdfThumbnail.generate(pdfThumbnail.loadConfig(filePath.toString()));
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
