/**
 * Unlicensed code created by A Softer Space, 2018
 * www.asofterspace.com/licenses/unlicense.txt
 */
package com.asofterspace.xdcReportCreator;

import com.asofterspace.toolbox.io.BinaryFile;
import com.asofterspace.toolbox.io.Directory;
import com.asofterspace.toolbox.io.File;
import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.JSON;
import com.asofterspace.toolbox.io.JsonFile;
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

	public final static String PROGRAM_TITLE = "XDC Report Creator";
	public final static String VERSION_NUMBER = "0.0.0.6(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "19. October 2018 - 20. December 2018";

	private final static String filesPath = "files/";
	private final static String inputPath = "input.json";
	private final static String templateFileName = filesPath + "template.xlsx";
	private final static String reportFileName = "report.xlsx";
	private final static String reportMacroFileName = "report.xlsm";
	private final static String reportPDFFileName = "report.pdf";
	private final static String macroFileName = filesPath + "vbaProject.bin";
	
	private static JSON inputData;
	
	private static XlsxFile template;
	
	private static XlsmFile macroTemplate;
	
	
	public static void main(String[] args) {
	
		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);
		
		System.out.println(Utils.getFullProgramIdentifierWithDate());

		cleanEnvironment();
		
		loadInputData();
		
		createTemplate();
		
		createMacroTemplate();
		
		// openTemplate();

		closeAllTemplates();
		
		System.out.println("Done!");
		System.out.println("A softer space wishes you a shiny day. :)");
	}
	
	private static void cleanEnvironment() {
	
		System.out.println("Cleaning environment after last run...");
		
		IoUtils.cleanAllWorkDirs();

		(new File(reportFileName)).delete();
		(new File(reportMacroFileName)).delete();
		// (new File(reportPDFFileName)).delete();
	}
	
	private static void loadInputData() {
	
		System.out.println("Loading the input data for the report from " + inputPath + "...");
		
		JsonFile inputFile = new JsonFile(inputPath);
		inputData = inputFile.getAllContents();
	}
	
	private static void createTemplate() {
	
		System.out.println("Loading the template...");
		
		JSON standardXDC = inputData.get("StandardXDC");
		
		template = new XlsxFile(templateFileName);

		List<XlsxSheet> sheets = template.getSheets();

		System.out.println("Adjusting the template to generate this particular report...");

		String footerdate = inputData.getString("Datum");

		if ((footerdate == null) || footerdate.equals("") || footerdate.equals("heute")) {
			DateFormat df = new SimpleDateFormat("dd.MM.yyyy");
			footerdate = df.format(new Date());
		}
		
		int pagenum = 0;
		
		for (XlsxSheet sheet : sheets) {
			// System.out.println("Adjusting sheet: " + sheet.getTitle());

			// Auftrag
			if (sheet.getTitle().startsWith("1. ")) {
				// System.out.println("Setting Angebotsnummer to " + inputData.getString("Angebotsnummer"));
				// System.out.println("(The old one was " + 			sheet.getCellContent("C9") + ")");
				sheet.setCellContent("C9", inputData.getString("Angebotsnummer"));
				
				// delete the scenario-based XDC for now
				// TODO :: actually read scenarios from input.json!
				sheet.deleteCellBlock("A20", "H28");
				
				sheet.setCellContent("C16", standardXDC.getString("Beschreibung"));
			}
			
			// Daten Standard XDC
			if (sheet.getTitle().startsWith("2. ")) {
				sheet.setCellContent("C8", inputData.getInteger("Berichtsjahr"));
				sheet.setCellContent("C10", inputData.getString("Firma"));
				sheet.setCellContent("C12", inputData.getString("FirmenISIN"));
				
				sheet.setCellContent("A23", inputData.getString("Firma"));
				sheet.setCellContent("A26", standardXDC.getString("DatenDatum"));
				
				sheet.setCellContent("A16", standardXDC.getLong("GHGScope1"));
				sheet.setCellContent("B16", standardXDC.getLong("GHGScope2"));
				sheet.setCellContent("C16", standardXDC.getLong("GHGScope3"));
				sheet.setCellContent("D16", standardXDC.getLong("EBITDA"));
				sheet.setCellContent("F16", standardXDC.getLong("Personalkosten"));
				
				Integer basisjahr = standardXDC.getInteger("Basisjahr");
				sheet.setCellContent("C36", basisjahr);
				sheet.setCellContent("C38", basisjahr);
				sheet.setCellContent("C40", basisjahr);
				sheet.setCellContent("C42", basisjahr);
				sheet.setCellContent("E34", basisjahr + "-2050");
				sheet.setCellContent("F34", basisjahr + "-2050");
				
				String entwicklungProcPA = standardXDC.getString("EntwicklungProcPA");
				sheet.setCellContent("E36", entwicklungProcPA + "% p.a.");
				sheet.setCellContent("E38", entwicklungProcPA + "% p.a.");
				sheet.setCellContent("F40", entwicklungProcPA + "% p.a.");
			}

			// Daten Szenariobasierte XDC
			if (sheet.getTitle().startsWith("3. ")) {
				// delete the scenario-based XDC for now
				// TODO :: actually read scenarios from input.json!
				sheet.deleteCellBlock("A1", "G60");
			}
			
			// Ergebnisse
			if (sheet.getTitle().startsWith("4. ")) {
				// delete the scenario-based XDC for now
				// TODO :: actually read scenarios from input.json!
				sheet.deleteCellBlock("A10", "H12");
				sheet.deleteCellBlock("A15", "K70");
			}
			
			// Berechnung
			if (sheet.getTitle().startsWith("5. ")) {
				sheet.setCellContent("B13", inputData.getString("Firma"));
				
				// delete the scenario-based XDC for now
				// TODO :: actually read scenarios from input.json!
				sheet.deleteCellBlock("A17", "AJ66");
			}
			
			// only adjust the footer for sheets that have a footer in the template
			if (sheet.hasFooter()) {
				// set the date in the footer
				sheet.setFooterContent("R", footerdate);
				
				/*
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
				*/
				// ACTUALLY, just set the footer to P (pagenum) and export everything as one big PDF via VBA! (see the part where we save as XLSM ^^)
				sheet.setFooterContent("C", "&P");
			}
			
			sheet.save();
		}
		
		System.out.println("Saving the report to " + reportFileName + "...");
		
		template.saveTo(reportFileName);
	}
	
	private static void createMacroTemplate() {
	
		System.out.println("Creating macro-enabled report at " + reportMacroFileName + "...");
		
		macroTemplate = template.convertToXlsm();
		
		/*
		// add the following Macro to the Workbook on Open:

		Private Sub Workbook_Open()

			'Create an array that contains all the sheets which should be exported to PDF or printed
			Dim sheetsToBePrinted() As String
			Dim workshAmount As Integer
			Dim worksh As Variant

			workshAmount = 0

			For Each worksh In Worksheets
				'Exclude sheet 5 - Berechnung (as it is basically the backend)
				If Not Left(worksh.Name, 1) = "5" Then
					workshAmount = workshAmount + 1
					ReDim Preserve sheetsToBePrinted(1 To workshAmount)
					sheetsToBePrinted(workshAmount) = worksh.Name
				End If
			Next
			
			'Activate and select all sheets which should be exported
			For Each worksh In sheetsToBePrinted
				Sheets(worksh).Activate
				ActiveSheet.UsedRange.Select
			Next

			'Actually export the sheets as PDF
			Sheets(sheetsToBePrinted).Select
			ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
				ThisWorkbook.Path & "/report.pdf", Quality:=xlQualityStandard, _
				IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

		End Sub
		*/
		
		// do NOT add the macro here, as we would need to also mess with binary files which we do not fully understand...
		// ugh!
		//
		// File exportSheetsMacro = new File(macroFileName);
		//
		// macroTemplate.addMacro(exportSheetsMacro);

		macroTemplate.saveTo(reportMacroFileName);
	}

	private static void openTemplate() {
	
		System.out.println("Opening the report in Excel..");
		
		String reportMacroAbsoluteFileName = (new File(reportMacroFileName)).getAbsoluteFilename();
		String fullExcelPath = inputData.getString("Excel");
		String excelExe = (new File(fullExcelPath)).getLocalFilename();
		String excelPath = fullExcelPath.substring(0, fullExcelPath.length() - excelExe.length() - 1);

		try {
			ProcessBuilder processBuilder = new ProcessBuilder(fullExcelPath, reportMacroAbsoluteFileName);
			processBuilder.directory(new java.io.File(excelPath));
			Process process = processBuilder.start();
		} catch (IOException e) {
			System.err.println("Oh no - Excel could not be started! Exception: " + e);
		}
	}
	
	private static void closeAllTemplates() {
	
		// close the files - this is important als XLSX and XLSM are zipped file types, and some temp files have to be cleared up!
		
		if (template != null) {
			template.close();
		}
		
		if (macroTemplate != null) {
			macroTemplate.close();
		}
	}

}
