package com.asofterspace.xdcReportCreator;

import com.asofterspace.toolbox.Utils;


public class Main {

	public final static String PROGRAM_TITLE = "XDC Report Creator (Java Part)";
	public final static String VERSION_NUMBER = "0.0.0.1(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "19. October 2018";

	public static void main(String[] args) {
	
		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);
		
		System.out.println(Utils.getFullProgramIdentifierWithDate());
	}

}
