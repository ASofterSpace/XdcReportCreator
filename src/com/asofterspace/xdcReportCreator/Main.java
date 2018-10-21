package com.asofterspace.xdcReportCreator;

import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.XlsxFile;
import com.asofterspace.toolbox.io.XlsxSheet;
import com.asofterspace.toolbox.Utils;

import java.util.List;


public class Main {

	public final static String PROGRAM_TITLE = "XDC Report Creator (Java Part)";
	public final static String VERSION_NUMBER = "0.0.0.1(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "19. October 2018";

	public static void main(String[] args) {
	
		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);
		
		IoUtils.cleanAllWorkDirs();

		System.out.println(Utils.getFullProgramIdentifierWithDate());

		XlsxFile template = new XlsxFile("template.xlsx");

		List<XlsxSheet> sheets = template.getSheets();

		for (XlsxSheet sheet : sheets) {
			System.out.println("Found sheet: " + sheet.getTitle());
		}
		
		template.saveTo("report.xlsx");
	}

}
