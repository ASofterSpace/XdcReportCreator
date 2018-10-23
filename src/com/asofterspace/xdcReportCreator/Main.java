package com.asofterspace.xdcReportCreator;

import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.JsonFile;
import com.asofterspace.toolbox.io.XlsxFile;
import com.asofterspace.toolbox.io.XlsxSheet;
import com.asofterspace.toolbox.Utils;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;


public class Main {

	public final static String PROGRAM_TITLE = "XDC Report Creator (Java Part)";
	public final static String VERSION_NUMBER = "0.0.0.2(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "19. October 2018 - 21. October 2018";

	public static void main(String[] args) {
	
		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);
		
		IoUtils.cleanAllWorkDirs();

		System.out.println(Utils.getFullProgramIdentifierWithDate());

		String inputPath = "input.json";
		
		System.out.println("Loading the input data for the report from " + inputPath + "...");
		
		JsonFile inputData = new JsonFile(inputPath);
		
		System.out.println("Loading the template...");
		
		XlsxFile template = new XlsxFile("template.xlsx");

		List<XlsxSheet> sheets = template.getSheets();

		System.out.println("Adjusting the template to generate this particular report...");

		String footerdate = inputData.getValue("Datum");

		if ((footerdate == null) || footerdate.equals("") || footerdate.equals("heute")) {
			DateFormat df = new SimpleDateFormat("dd.MM.yyyy");
			footerdate = df.format(new Date());
		}
		
		int pagenum = 0;
		
		for (XlsxSheet sheet : sheets) {
			// System.out.println("Adjusting sheet: " + sheet.getTitle());

			// Auftrag
			if (sheet.getTitle().startsWith("1. ")) {
				// System.out.println("Setting Angebotsnummer to " + inputData.getValue("Angebotsnummer"));
				// System.out.println("(The old one was " + 			sheet.getCellContent("C9") + ")");
				sheet.setCellContent("C9", inputData.getValue("Angebotsnummer"));
			}

			// only adjust the footer for sheets that have a footer in the template
			if (sheet.hasFooter()) {
				// set the date in the footer
				sheet.setFooterContent("R", footerdate);
				
				// set the page number in the footer
				sheet.setFooterContent("C", "&P+"+pagenum);
				// TODO :: what about multi-page sheets? (
				// we need to count up plus the amount of pages for the particular worksheet...
				// does Excel have a funky function for that?
				pagenum++;
				if (sheet.getTitle().startsWith("3. ")) {
					pagenum++;
				}
				if (sheet.getTitle().startsWith("4. ")) {
					pagenum += 2;
				}
				if (sheet.getTitle().startsWith("6. ")) {
					pagenum += 4;
				}
			}
			
			sheet.save();
		}
		
		String reportFileName = "report.xlsx";
		
		System.out.println("Saving the report to " + reportFileName + "...");
		
		template.saveTo(reportFileName);
		
		template.close();
		
		System.out.println("Done!");
		System.out.println("A softer space wishes you a shiny day. :)");
	}

}
