package com.org.thumbnail;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

import org.jodconverter.LocalConverter;
import org.jodconverter.LocalConverter.Builder;
import org.jodconverter.filter.text.PageCounterFilter;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;

public class DocumentThumbnailGenerator {

	public void generate(Properties prop) {
		String filePath = prop.getProperty("SorceFile");
		String thumbnailPath = prop.getProperty("OutputFolder");
		String imagePrefix = prop.getProperty("Format");

		File inputFile = new File(filePath);

		File newFolder = new File(thumbnailPath);
		boolean created = newFolder.mkdirs();

		int count = 1;
		try {
			int page = 1;

			while (count > 0) {
				count--;
				Builder builder = LocalConverter.builder();
				File outputFile = new File(thumbnailPath + "page_" + page + imagePrefix);
				PageCounterFilter pageCountfilter = new PageCounterFilter();
				CustomPageSelectorFilter pageSelectfilters = new CustomPageSelectorFilter(page);
				builder = builder.filterChain(pageSelectfilters);
				if (page == 1) {
					builder.filterChain(pageCountfilter).build().convert(inputFile).to(outputFile).execute();
					count = pageCountfilter.getPageCount() - 1;
				} else {
					builder.build().convert(inputFile).to(outputFile).execute();
				}
				page++;
			}
		} catch (OfficeException e) {
			System.out.println("officeException caught here:::" + e.getMessage());
		} catch (Exception e) {
			System.out.println("exception caught from here:::" + e.getMessage());
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

		DocumentThumbnailGenerator docThumbnail = new DocumentThumbnailGenerator();
		final LocalOfficeManager officeManager = LocalOfficeManager.builder().officeHome(getOfficeHome()).install()
				.build();
		try {
			officeManager.start();
			docThumbnail.generate(docThumbnail.loadConfig(filePath.toString()));
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
	 * "/Applications/LibreOffice.app/Contents" for linux -> "/opt/libreoffice"
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
