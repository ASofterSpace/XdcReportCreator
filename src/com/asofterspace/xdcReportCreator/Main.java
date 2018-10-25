package com.asofterspace.xdcReportCreator;

import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.JSON;
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
		
		JsonFile inputFile = new JsonFile(inputPath);
		JSON inputData = inputFile.getAllContents();
		JSON standardXDC = inputData.get("StandardXDC");
		
		System.out.println("Loading the template...");
		
		XlsxFile template = new XlsxFile("template.xlsx");

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
				// ACTUALLY, just set the footer to P (pagenum) and export everything as one big PDF via VBA!
				sheet.setFooterContent("C", "&P");
				/*
				for that, we can use the following VBA:
				
Private Sub Workbook_Open()

    For i = 1 To 6
        Sheets(i).Activate
        ActiveSheet.UsedRange.Select
    Next i

    ThisWorkbook.Sheets(Array(1, 2, 3, 4, 5, 6)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ThisWorkbook.Path & "/report.pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
      
End Sub

				however, we also have to convert the file to xlsm, AND we have to specify the working path such that the file is stored there
				*/
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
