package com.paulo;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLProperties.CustomProperties;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;

import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.GregorianCalendar;
import java.util.LinkedList;
import java.util.Queue;

public class TemplateWriter {
    public static void createTemplate(String path) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(10);
        setWorkbookProperties(workbook);
        setTemplateSheet(workbook);
        setDataSheet(workbook);


        SXSSFWriter.createFile(path, workbook);
    }

    private static void setWorkbookProperties(SXSSFWorkbook workbook) {
        POIXMLProperties xmlProperties = workbook.getXSSFWorkbook().getProperties();

        CoreProperties coreProperties = xmlProperties.getCoreProperties();
        coreProperties.setKeywords("CorePropertyKeyword");

        CustomProperties customProperties = xmlProperties.getCustomProperties();
        customProperties.addProperty("customPropertyKey", "customPropertyValue");
    }

    private static void setTemplateSheet(SXSSFWorkbook workbook) {
        XSSFWorkbook wb = workbook.getXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet("TemplateSheet");
        SXSSFRow row = sheet.createRow(0);

        // see org.apache.poi.ss.usermodel.BuiltinFormats
        // NUMERIC TYPES
        XSSFCellStyle columnStyle = wb.createCellStyle();
        XSSFDataFormat dataFormat = wb.createDataFormat();
        columnStyle.setDataFormat(dataFormat.getFormat("#.##0"));
        sheet.setDefaultColumnStyle(0, columnStyle);
        row.createCell(0).setCellValue(3421654.987421);

        // DATE TYPES
        columnStyle = wb.createCellStyle();
        dataFormat = wb.createDataFormat();
        columnStyle.setDataFormat(dataFormat.getFormat("m/d/yy h:mm"));
        sheet.setDefaultColumnStyle(1, columnStyle);
        SXSSFCell cell = row.createCell(1);
        cell.setCellValue(now());


        // Liste déroulante
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("nomDeMaContrainteSansTiretEtSansCaracteresSpeciaux");
        CellRangeAddressList addressList = new CellRangeAddressList(0, SpreadsheetVersion.EXCEL2007.getMaxRows() - 1, 2, 2);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
        validation.createPromptBox("MCM Team", "Nom des personnes dans la team MCI/MCM");
        validation.setShowPromptBox(true);
        sheet.addValidationData(validation);

        // font and whatever (style)
        XSSFFont font = new XSSFFont(CTFont.Factory.newInstance());
        font.setBold(true);
        font.setColor(HSSFColor.WHITE.index);
        font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        font.registerTo(wb.getStylesSource());

        XSSFCellStyle cellStyle = workbook.getXSSFWorkbook().createCellStyle();
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 112, (byte) 187, (byte) 50}));
        SXSSFCell cellFont = row.createCell(3);
        cellFont.setCellValue("FONT THINGS");
        cellFont.setCellStyle(cellStyle);
    }

    private static GregorianCalendar now() {
        return GregorianCalendar.from(ZonedDateTime.ofInstant(Instant.now(), ZoneId.systemDefault()));
    }

    private static void setDataSheet(SXSSFWorkbook workbook) {
        Queue<String> mciTeam = new LinkedList<>();
        mciTeam.add("Ste");
        mciTeam.add("Ronron");
        mciTeam.add("Jéwôme");
        mciTeam.add("SuperTimor");
        mciTeam.add("Pauliméro");
        mciTeam.add("Bazou");
        mciTeam.add("Bastieng");
        mciTeam.add("Poompoom");
        mciTeam.add("Cissou");
        mciTeam.add("Lapeyro");

        SXSSFSheet sheet = workbook.createSheet("DataSheet");
        SXSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("La team MCM");
        for (int i = 1; i <= 10; i++) {
            sheet.createRow(i).createCell(0).setCellValue(mciTeam.poll());
        }

        Name validationFieldsName = workbook.createName();
        validationFieldsName.setNameName("nomDeMaContrainteSansTiretEtSansCaracteresSpeciaux");
        validationFieldsName.setRefersToFormula(sheet.getSheetName() + "!$A$2:$A$11");

    }
}
