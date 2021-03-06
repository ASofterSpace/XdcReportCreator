/**
 * Unlicensed code created by A Softer Space, 2018
 * www.asofterspace.com/licenses/unlicense.txt
 */
package com.asofterspace.xdcReportImprover;

import com.asofterspace.toolbox.io.BinaryFile;
import com.asofterspace.toolbox.io.Directory;
import com.asofterspace.toolbox.io.File;
import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.JSON;
import com.asofterspace.toolbox.io.JsonFile;
import com.asofterspace.toolbox.io.PdfFile;
import com.asofterspace.toolbox.io.PdfObject;
import com.asofterspace.toolbox.io.XlsmFile;
import com.asofterspace.toolbox.io.XlsxFile;
import com.asofterspace.toolbox.io.XlsxSheet;
import com.asofterspace.toolbox.Utils;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;


public class Main {

	public final static String PROGRAM_TITLE = "XDC Report Improver";
	public final static String VERSION_NUMBER = "0.0.0.1(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "20. December 2018";

	private final static String filesPath = "files/";
	private final static String reportPDFFileName = "report.pdf";
	private final static String reportPDFHighQualityFileName = "report_high_quality.pdf";
	
	private static JSON inputData;
	
	private static XlsxFile template;
	
	private static XlsmFile macroTemplate;
	
	
	public static void main(String[] args) {
	
		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);
		
		System.out.println(Utils.getFullProgramIdentifierWithDate());

		replacePicsInPdf(reportPDFFileName, reportPDFHighQualityFileName);

		System.out.println("Done!");
		System.out.println("A softer space wishes you a shiny day. :)");
	}
	
	private static void replacePicsInPdf(String origPdfPath, String newPdfPath) {

		System.out.println("Loading the report generated by Excel from " + origPdfPath + "...");
		
		File oldFile = new File(origPdfPath);
		oldFile.copyToDisk(newPdfPath);
	
		PdfFile pdf = new PdfFile(newPdfPath);
		List<PdfObject> objects = pdf.getObjects();
		
		System.out.println("Loading the high quality logos from " + filesPath + "...");
		
		String rightPath = filesPath + "right.jpg";
		BinaryFile rightPicFile = new BinaryFile(rightPath);
		String rightPicContent = rightPicFile.loadContentStr();

		String xdcPath = filesPath + "xdc.jpg";
		BinaryFile xdcPicFile = new BinaryFile(xdcPath);
		String xdcPicContent = xdcPicFile.loadContentStr();

		for (PdfObject obj : objects) {
			if ("/XObject".equals(obj.getDictValue("/Type"))) {
				if ("/Image".equals(obj.getDictValue("/Subtype"))) {
					obj.setDictValue("/ColorSpace", "/DeviceRGB");
					obj.setDictValue("/BitsPerComponent", "8");
					obj.setDictValue("/Filter", "/DCTDecode");
					obj.setDictValue("/Interpolate", "true");
					obj.removeDictValue("/SMask");
					obj.removeDictValue("/Matte");
					
					// the small right picture in the original report has a width of 272, the xdc picture has a width of 281
					boolean replaceWithRightPic = obj.getDictValue("/Width").startsWith("27");
					
					if (replaceWithRightPic) {
						obj.setDictValue("/Width", "2540");
						obj.setDictValue("/Height", "794");
						obj.setStreamContent(rightPicContent);
					} else {
						obj.setDictValue("/Width", "2717");
						obj.setDictValue("/Height", "883");
						obj.setStreamContent(xdcPicContent);
					}
				}
			}
		}
		
		System.out.println("Saving the resulting report as " + newPdfPath + "...");
		
		pdf.save();
	}

}
