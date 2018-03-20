/*
 * Copyright 2016 Telstra Corporation.
 *
 * This program is the property of Telstra Corporation.
 * You may not use this code without the express permission of
 * the designated business owner of the software in Telstra.
 * Parts of this program are copyright Dialog Information Technology
 * and licensed for the use of Telstra Corporation as part of this
 * software.
 */
package canrad.misc;

import dialog.excel.ExcelStyleGenerator;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 *
 * @author SeetoB
 */
public class TIMReportGenerator
{
    public static final String EXCEL_SHEET_NAME = "F03 Tier-2";
    public static final String FONT_NAME_CALIBRI = "Calibri";

    public static final String WORKSHEET_HEADING = "Worksheet 3 / Section A";
    public static final String EQUIP_DETAIL_HEADING = "CANRAD SITE and EQUIPMENT DETAILS";
    public static final String PRINT_INST_HEADING = "Print document A4 Landscape / Margins Top & Bott. 1.4cm  \n" + "L & R 1cm / Header & Footer 0.8cm ";
    public static final String EQUIP_DETAIL_SUBHEADING = "CANRAD EXISTING EQUIPMENT DETAILS";
    public static final String STRUCTURE_DETAIL_HEADING = "Structure Detail";
    public static final String STRUCTURE_GUY_WIRES_HEADING = "Structure Guy Wires  (010623 Iss. 12 / Sect. 4.12.10)";
    public static final String GUY_TENSIONS_HEADING = "Guy Tensions";
    public static final String ANTENNA_DETAIL_HEADING = "Antenna Detail";
    public static final String ANTENNA_MOUNTING_HEIGHT_HEADING = "Antenna Mounting Height";
    public static final String FEEDER_DETAIL_HEADING = "Feeder Detail";
    public static final String SHELTER_HEADING = "Shelter";
    public static final String SHELTER_NUMBER_HEADING = "Shelter Number";
    public static final String SHELTER_TYPE_HEADING = "Shelter Type";
    public static final String TRANSCEIVER_HEADING = "Transceiver";
    public static final String TRANSCEIVER_NUMBER_HEADING = "Transceiver Number";
    public static final String TRANSCEIVER_TYPE_HEADING = "Transceiver Type Code";
    public static final String BATTERY_BANK_HEADING = "Battery Bank";
    public static final String POWER_CONVERTER_HEADING = "Power Converter";
    public static final String PRIMARY_POWER_SOURCE_HEADING = "Primary Power Source";

    public static final String SUMMARY = "Below is a list of Telstra's radio equipment that is on site according to CANRAD. Verify if CANRAD equipment records are correct; Check if CANRAD aligns with the RF equipment that is actually on site. Record any variations in Section B: CANRAD EQUIPMENT VARIATION DETAILS.";

    public static final String SITE_NAME_LABEL = "Site Name:";
    public static final String SITE_ID_LABEL = "CANRAD Site ID:";
    public static final String DOWNLOAD_DATE_LABEL = "Data Download date:";
    public static final String INSP_1_LABEL = "Inspector 1:";
    public static final String MOBILE_1_LABEL = "Mobile:";
    public static final String INSP_2_LABEL = "Inspector 2:";
    public static final String MOBILE_2_LABEL = "Mobile:";
    public static final String SITE_COORD_LABEL = "Site Coordinates";
    public static final String MAP_DATUM_LABEL = "Map Datum";
    public static final String COORD_PRECISION_LABEL = "Coord. Precision";
    public static final String ZONE_LABEL = "Zone";
    public static final String EASTING_LABEL = "Easting";
    public static final String NORTHING_LABEL = "Northing";
    public static final String SITE_RL_LABEL = "Site RL (m)";
    public static final String STRUCTURE_NUMBER_LABEL = "Structure Number";
    public static final String TYPE_LABEL = "Type";
    public static final String HEIGHT_LABEL = "Height";
    public static final String EXTENSION_LABEL = "Extension";
    public static final String PAINT_DETAILS_LABEL = "Paint Details";
    public static final String LIGHTS_LABEL = "Lights";
    public static final String ANTICLIMB_LABEL = "Anticlimb Dvc.";
    public static final String OWNER_LABEL = "Owner";
    public static final String STATUS_LABEL = "Status";
    public static final String NOTES_LABEL = "Notes";
    public static final String GUY_CABLE_LABEL = "Guy Cable Number";
    public static final String ATTACH_HEIGHT_LABEL = "Attach Height";
    public static final String FACES_CORNER_LABEL = "Faces / Corners";
    public static final String BEARING_LABEL = "Bearing";
    public static final String KN_DESIGNED_LABEL = "kN Designed";
    public static final String KN_MEASURED_LABEL = "kN Measured";
    public static final String CABLE_TYPE_LABEL = "Cable Type / Construction Details";
    public static final String TORQUE_STAB_LABEL = "Torque Stab Y/N";
    public static final String ANTENNA_NUMBER_LABEL = "Antenna Number";
    public static final String CANTILEVER_LABEL = "Ant. Attached Detail \n" + "e.g. Cantilever";
    public static final String ANTENNA_TYPE_LABEL = "Antenna Type";
    public static final String DEST_BEARING_LABEL = "Dest. Bearing (T)";
    public static final String POLARITY_LABEL = "Polarity";
    public static final String FACE_CNR_LABEL = "Face / Corner";
    public static final String MOUNT_HEIGHT_LABEL = "Mount Ht.";
    public static final String ANTENNA_PT_LABEL = "Ant. Mid Point";
    public static final String STRUCT_NO_LABEL = "Struct. No";
    public static final String FEEDER_NO_LABEL = "Feeder Number";
    public static final String FEEDER_TYPE_LABEL = "Feeder Type";
    public static final String FEEDER_LENGTH_LABEL = "Feeder Length";
    public static final String A_END_LABEL = "A End (Antenna Number)";
    public static final String B_END_LABEL = "B END (Transceiver Notes)";
    public static final String BATTERY_BANK_LABEL = "Battery Bank no";
    public static final String BATTERY_TYPE_LABEL = "Battery Type";
    public static final String NUM_BATTERIES_LABEL = "No of Batteries";
    public static final String BATTERY_BANK_VOLT_LABEL = "Battery Bank Voltage";
    public static final String SHELTER_NO_LABEL = "Shelter No.";
    public static final String CONVERTER_NO_LABEL = "Converter Number";
    public static final String CONVERTER_TYPE_LABEL = "Power Converter Type";
    public static final String POWER_SOURCE_LABEL = "Power Source (Solar/Generator/Mains)";
    public static final String POWER_SOURCE_TYPE_LABEL = "Primary Power Source Type";
    public static final String NUM_POWER_SOURCE_LABEL = "Number of Primary Power sources";
    public static final String SITE_VOLTAGE_LABEL = "Site Voltage";
    public static final String SERIAL_NUMBER_LABEL = "Serial Numbers";

    public static final String STYLE1 = "STYLE1";
    public static final String STYLE2 = "STYLE2";
    public static final String STYLE3 = "STYLE3";
    public static final String STYLE4 = "STYLE4";
    public static final String STYLE5 = "STYLE5";
    public static final String STYLE6 = "STYLE6";
    public static final String STYLE7 = "STYLE7";
    public static final String STYLE8 = "STYLE8";
    public static final String STYLE9 = "STYLE9";
    public static final String STYLE10 = "STYLE10";
    public static final String STYLE11 = "STYLE11";
    
    private final HSSFWorkbook wb;
    private final Sheet sheet;
    private final CreationHelper createHelper;
    private final Map<String, CellStyle> styleMap;
    private int rowNum = 0;

    public TIMReportGenerator()
    {
        wb = new HSSFWorkbook();
        sheet = wb.createSheet(EXCEL_SHEET_NAME);
        createHelper = wb.getCreationHelper();
        styleMap = new HashMap<>();
    }

    public void setupReportStyles()
    {
        setupStyles();
        resetColourPalette();
    }

    public void createMainHeader(List<String> dataList)
    {
        createHeader();
        createSummaryHdr();
        createSiteHdr(dataList);
        createInspHdr();
    }

    public void writeToSpreadsheet(FileOutputStream fileOut)
    {
        try
        {
            wb.write(fileOut);
        } catch (FileNotFoundException ex)
        {
            Logger.getLogger(TIMReportGenerator.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex)
        {
            Logger.getLogger(TIMReportGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void setupStyles()
    {
        // CREATE FONTS **********
        Font font1 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 12, true, IndexedColors.WHITE.getIndex());
        Font font2 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 8, false, IndexedColors.WHITE.getIndex());
        Font font3 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 10, true, IndexedColors.BLACK.getIndex());
        Font font4 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 10, false, IndexedColors.BLACK.getIndex());
        Font font5 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 10, true, IndexedColors.WHITE.getIndex());
        //font6 is used to display equipment details
        Font font6 = ExcelStyleGenerator.createFont(wb, FONT_NAME_CALIBRI, (short) 8, false, IndexedColors.BLACK.getIndex());

        //CREATE STYLES ***********
        CellStyle style1 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.TOP, HorizontalAlignment.LEFT, IndexedColors.LIGHT_BLUE.getIndex());
        style1.setFont(font1);
        styleMap.put(STYLE1, style1);
        CellStyle style2 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.TOP, HorizontalAlignment.CENTER, IndexedColors.LIGHT_BLUE.getIndex());
        style2.setFont(font1);
        styleMap.put(STYLE2, style2);
        CellStyle style3 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.TOP, HorizontalAlignment.CENTER, IndexedColors.LIGHT_BLUE.getIndex());
        style3.setFont(font2);
        styleMap.put(STYLE3, style3);
        CellStyle style4 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.TOP, HorizontalAlignment.LEFT, IndexedColors.GREY_25_PERCENT.getIndex());
        style4.setFont(font3);
        styleMap.put(STYLE4, style4);
        CellStyle style5 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.LEFT, IndexedColors.GREY_40_PERCENT.getIndex());
        style5.setFont(font3);
        styleMap.put(STYLE5, style5);
        CellStyle style6 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.CENTER, IndexedColors.WHITE.getIndex());
        style6.setFont(font4);
        styleMap.put(STYLE6, style6);
        CellStyle style7 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.CENTER, IndexedColors.LIGHT_BLUE.getIndex());
        style7.setFont(font5);
        styleMap.put(STYLE7, style7);
        CellStyle style8 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.CENTER, IndexedColors.GREY_40_PERCENT.getIndex());
        style8.setFont(font3);
        styleMap.put(STYLE8, style8);
        CellStyle style9 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.LEFT, IndexedColors.LIGHT_BLUE.getIndex());
        style9.setFont(font5);
        styleMap.put(STYLE9, style9);
        CellStyle style10 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.CENTER, IndexedColors.GREY_50_PERCENT.getIndex());
        style10.setFont(font5);
        styleMap.put(STYLE10, style10);
        //This style to write the variable data out
        CellStyle style11 = ExcelStyleGenerator.createStyle(wb, VerticalAlignment.CENTER, HorizontalAlignment.CENTER, IndexedColors.WHITE.getIndex());
        style11.setFont(font6);
        styleMap.put(STYLE11, style11);
    }

    private void resetColourPalette()
    {
        // Customise colours
        //Customise Grey_25_Percent
        ExcelStyleGenerator.resetColourPalette(wb.getCustomPalette(), HSSFColor.GREY_25_PERCENT.index, (byte) 242, (byte) 242, (byte) 242);

        //Customise Grey_40_Percent
        ExcelStyleGenerator.resetColourPalette(wb.getCustomPalette(), HSSFColor.GREY_40_PERCENT.index, (byte) 217, (byte) 217, (byte) 217);

        //Customise Light_Blue
        ExcelStyleGenerator.resetColourPalette(wb.getCustomPalette(), HSSFColor.LIGHT_BLUE.index, (byte) 79, (byte) 129, (byte) 189);
    }

    private void createHeader()
    {
        //Create header on row 1 - rowNum should be 0
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 2, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);

        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 600);
        ExcelStyleGenerator.createCell(0, 2, styleMap.get(STYLE1), row);
        ExcelStyleGenerator.createCell(2, 7, styleMap.get(STYLE2), row);
        ExcelStyleGenerator.createCell(7, 11, styleMap.get(STYLE3), row);

        ExcelStyleGenerator.populateRow(row, WORKSHEET_HEADING, createHelper, 0);
        ExcelStyleGenerator.populateRow(row, EQUIP_DETAIL_HEADING, createHelper, 2);
        ExcelStyleGenerator.populateRow(row, PRINT_INST_HEADING, createHelper, 7);
    }

    private void createSummaryHdr()
    {
        //Create summary row on row 2 - rowNum should be 1
        rowNum++;
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 600);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE4), row);
        ExcelStyleGenerator.populateRow(row, SUMMARY, createHelper, 0);
        
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 10);
    }

    private void createSiteHdr(List<String> dataList)
    {
        //rowNum should be 2
        rowNum++;
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 400);
        ExcelStyleGenerator.createCell(0, 1, styleMap.get(STYLE5), row);

        ExcelStyleGenerator.createCell(1, 3, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(3, 4, styleMap.get(STYLE5), row);
        ExcelStyleGenerator.createCell(4, 6, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(6, 8, styleMap.get(STYLE5), row);
        ExcelStyleGenerator.createCell(8, 10, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(10, 11, styleMap.get(STYLE5), row);

        //Populate variable data
        ExcelStyleGenerator.populateRow(row, SITE_NAME_LABEL, createHelper, 0);
        ExcelStyleGenerator.populateRow(row, dataList.get(0), createHelper, 1);
        ExcelStyleGenerator.populateRow(row, SITE_ID_LABEL, createHelper, 3);
        ExcelStyleGenerator.populateRow(row, dataList.get(1), createHelper, 4);
        ExcelStyleGenerator.populateRow(row, DOWNLOAD_DATE_LABEL, createHelper, 6);
        
        String checkoutTimestamp = "";
        if (!dataList.get(2).equals(""))
        {
            ZonedDateTime result = ZonedDateTime.parse(dataList.get(2), DateTimeFormatter.ISO_DATE_TIME);
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm a");
            checkoutTimestamp = result.toLocalDateTime().format(formatter);
        }
        ExcelStyleGenerator.populateRow(row, checkoutTimestamp, createHelper, 8);
        
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 2);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 7);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 8, 10);
    }

    private void createInspHdr()
    {
        //rowNum should be 3
        rowNum++;
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);

        ExcelStyleGenerator.createCell(0, 1, styleMap.get(STYLE5), row);

        //TODO populate variable data - dont forget to increment rowNum
        ExcelStyleGenerator.createCell(1, 2, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(2, 3, styleMap.get(STYLE5), row);
        ExcelStyleGenerator.createCell(3, 4, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(4, 6, styleMap.get(STYLE5), row);
        ExcelStyleGenerator.createCell(6, 8, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(8, 9, styleMap.get(STYLE5), row);
        ExcelStyleGenerator.createCell(9, 10, styleMap.get(STYLE6), row);
        ExcelStyleGenerator.createCell(10, 11, styleMap.get(STYLE5), row);

        ExcelStyleGenerator.populateRow(row, INSP_1_LABEL, createHelper, 0);
        ExcelStyleGenerator.populateRow(row, MOBILE_1_LABEL, createHelper, 2);
        ExcelStyleGenerator.populateRow(row, INSP_2_LABEL, createHelper, 4);
        ExcelStyleGenerator.populateRow(row, MOBILE_2_LABEL, createHelper, 8);
        
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 7);
    }

    public void createSiteCoordDetailHdr()
    {
        //rowNum should be 4
        rowNum++;
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE7), row);
        ExcelStyleGenerator.populateRow(row, EQUIP_DETAIL_SUBHEADING, createHelper, 0);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 10);

        //rowNum should be 5
        Row row2 = sheet.createRow((short) ++rowNum);
        row2.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row2);

        List<String> stringList = new ArrayList<>();
        stringList.add(SITE_COORD_LABEL);
        stringList.add(MAP_DATUM_LABEL);
        stringList.add(COORD_PRECISION_LABEL);
        stringList.add(ZONE_LABEL);
        stringList.add(EASTING_LABEL);
        stringList.add(NORTHING_LABEL);
        stringList.add(SITE_RL_LABEL);

        createSubHeading1(stringList, row2);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);
    }

    public void createSiteCoordDetailRow(List<String> dataList)
    {
        //rowNum should be 6
        rowNum++;
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE6), row);

        //Populate variable data
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);
        ExcelStyleGenerator.populateRow(row, dataList, createHelper); 
    }

    public void createStructureDetailHdr()
    {
        //rowNum should be 7
        rowNum++;
        createEquipHdr(STRUCTURE_DETAIL_HEADING);

        //Sub heading for Structure Detail ************************
        //rowNum should be 8
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 300);

        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(STRUCTURE_NUMBER_LABEL);
        stringList.add(TYPE_LABEL);
        stringList.add(HEIGHT_LABEL);
        stringList.add(EXTENSION_LABEL);
        stringList.add(PAINT_DETAILS_LABEL);
        stringList.add(LIGHTS_LABEL);
        stringList.add(ANTICLIMB_LABEL);
        stringList.add(OWNER_LABEL);
        stringList.add(STATUS_LABEL);
        stringList.add(NOTES_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);
    }

    public void createStructureDetailRow(List<String> dataList)
    {
        //Populate variable data
        Row structureRow = sheet.createRow((short) ++rowNum);
        structureRow.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), structureRow);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);
        ExcelStyleGenerator.populateRow(structureRow, dataList, createHelper); 
    }
    
    public void createStructureGuyWiresDetailHdr()
    {
        rowNum++;
        createEquipHdrWithGrey(STRUCTURE_GUY_WIRES_HEADING, GUY_TENSIONS_HEADING);

        //Sub heading for Structure Guy Wires *********************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(STRUCTURE_NUMBER_LABEL);
        stringList.add(GUY_CABLE_LABEL);
        stringList.add(ATTACH_HEIGHT_LABEL);
        stringList.add(FACES_CORNER_LABEL);
        stringList.add(STATUS_LABEL);
        stringList.add(BEARING_LABEL);
        stringList.add(KN_DESIGNED_LABEL);
        stringList.add(KN_MEASURED_LABEL);
        stringList.add(CABLE_TYPE_LABEL);
        stringList.add("");
        stringList.add(TORQUE_STAB_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 8, 9);
    }

    public void createStructureGuyWiresDetailRow(List<String> dataList)
    {
        //TODO Variable data - dont forget to increment rowNum
        Row row1 = sheet.createRow((short) ++rowNum);
        row1.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), row1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 8, 9);
        ExcelStyleGenerator.populateRow(row1, dataList, createHelper); 
    }

    public void createAntennaDetailHdr()
    {
        rowNum++;
        createEquipHdrWithGrey(ANTENNA_DETAIL_HEADING, ANTENNA_MOUNTING_HEIGHT_HEADING);

        //Sub heading Antenna Detail *******************************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(ANTENNA_NUMBER_LABEL);
        stringList.add(CANTILEVER_LABEL);
        stringList.add(ANTENNA_TYPE_LABEL);
        stringList.add(DEST_BEARING_LABEL);
        stringList.add(POLARITY_LABEL);
        stringList.add(FACE_CNR_LABEL);
        stringList.add(MOUNT_HEIGHT_LABEL);
        stringList.add(ANTENNA_PT_LABEL);
        stringList.add(STRUCT_NO_LABEL);
        stringList.add(OWNER_LABEL);
        stringList.add(STATUS_LABEL);
        createSubHeading1(stringList, row);
    }

    public void createAntennaDetailRow(List<String> dataList)
    {
        //Populate variable data
        Row antennaRow = sheet.createRow((short) ++rowNum);
        antennaRow.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), antennaRow);
        ExcelStyleGenerator.populateRow(antennaRow, dataList, createHelper); 
    }

    public void createFeederDetailHdr()
    {
        rowNum++;
        createEquipHdr(FEEDER_DETAIL_HEADING);

        //Sub heading for Feeder Detail ****************************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(FEEDER_NO_LABEL);
        stringList.add(FEEDER_TYPE_LABEL);
        stringList.add(FEEDER_LENGTH_LABEL);
        stringList.add(OWNER_LABEL);
        stringList.add(A_END_LABEL);
        stringList.add("");
        stringList.add(B_END_LABEL);
        stringList.add("");
        stringList.add(STATUS_LABEL);
        stringList.add(NOTES_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 7);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);
    }

    public void createFeederDetailRow(List<String> dataList)
    {
        //Populate variable data
        Row feederRow = sheet.createRow((short) ++rowNum);
        feederRow.setHeight((short) 600);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), feederRow);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 7);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);
        ExcelStyleGenerator.populateRow(feederRow, dataList, createHelper); 
    }

    public void createShelterDetailHdr()
    {
        rowNum++;
        createEquipHdr(SHELTER_HEADING);
        ++rowNum;
        createSubHeading2(SHELTER_NUMBER_HEADING, SHELTER_TYPE_HEADING);
    }

    public void createShelterDetailRow(List<String> dataList)
    {
        //Populate variable data
        Row shelterRow = sheet.createRow((short) ++rowNum);
        shelterRow.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), shelterRow);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 5, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);
        ExcelStyleGenerator.populateRow(shelterRow, dataList, createHelper); 
    }

    public void createTransceiverDetailHdr()
    {
        rowNum++;
        createEquipHdr(TRANSCEIVER_HEADING);
        ++rowNum;
        createSubHeading2(TRANSCEIVER_NUMBER_HEADING, TRANSCEIVER_TYPE_HEADING);
    }

    public void createTransDetailRow(List<String> dataList)
    {
        //Populate variable data
        Row transceiverRow = sheet.createRow((short) ++rowNum);
        transceiverRow.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), transceiverRow);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 5, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);
        ExcelStyleGenerator.populateRow(transceiverRow, dataList, createHelper); 
    }

    public void createBatteryBankDetailHdr()
    {
        rowNum++;
        createEquipHdr(BATTERY_BANK_HEADING);

        //Sub heading for Battery Bank ********************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(BATTERY_BANK_LABEL);
        stringList.add(BATTERY_TYPE_LABEL);
        stringList.add("");
        stringList.add(NUM_BATTERIES_LABEL);
        stringList.add(BATTERY_BANK_VOLT_LABEL);
        stringList.add("");
        stringList.add("");
        stringList.add(SHELTER_NO_LABEL);
        stringList.add(STATUS_LABEL);
        stringList.add(NOTES_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 2);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);

    }

    public void createBatteryBankDetailRow(List<String> dataList)
    {
        //TODO variable data - dont forget to increment rowNum
        Row row1 = sheet.createRow((short) ++rowNum);
        row1.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), row1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 2);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 9, 10);
        ExcelStyleGenerator.populateRow(row1, dataList, createHelper); 
    }

    public void createPowerConverterDetailHdr()
    {
        rowNum++;
        createEquipHdr(POWER_CONVERTER_HEADING);

        //Sub heading for Power Converter *************************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 400);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(CONVERTER_NO_LABEL);
        stringList.add(CONVERTER_TYPE_LABEL);
        stringList.add("");
        stringList.add("");
        stringList.add(STATUS_LABEL);
        stringList.add(SHELTER_NO_LABEL);
        stringList.add(NOTES_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 10);
    }

    public void createPowerConverterDetailRow(List<String> dataList)
    {
        //TODO variable data - dont forget to increment rowNum
        Row row1 = sheet.createRow((short) ++rowNum);
        row1.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), row1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 10);
        ExcelStyleGenerator.populateRow(row1, dataList, createHelper); 
    }

    public void createPrimaryPowerSourceDetailHdr()
    {
        rowNum++;
        createEquipHdr(PRIMARY_POWER_SOURCE_HEADING);

        //Sub heading for Primary Power Source *******************
        Row row = sheet.createRow((short) ++rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(POWER_SOURCE_LABEL);
        stringList.add("");
        stringList.add(POWER_SOURCE_TYPE_LABEL);
        stringList.add("");
        stringList.add(NUM_POWER_SOURCE_LABEL);
        stringList.add("");
        stringList.add(SITE_VOLTAGE_LABEL);
        stringList.add(STATUS_LABEL);
        stringList.add(SERIAL_NUMBER_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 2, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 8, 10);
    }

    public void createPrimaryPowerSourceDetailRow(List<String> dataList)
    {
        //variable data - dont forget to increment rowNum
        Row row1 = sheet.createRow((short) ++rowNum);
        row1.setHeight((short) 500);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE11), row1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 1);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 2, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 4, 5);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 8, 10);
        ExcelStyleGenerator.populateRow(row1, dataList, createHelper); 
    }

    public void autoSizeColumns()
    {
        int numberOfSheets = wb.getNumberOfSheets();

        for (int i = 0; i < numberOfSheets; i++)
        {
            Sheet sheet = wb.getSheetAt(i);
            for(int col = 0; col < 11; col++)
                sheet.autoSizeColumn(col);
            /*
            if (sheet.getPhysicalNumberOfRows() > 0)
            {
                Row row = sheet.getRow(0);
                sheet.getRow(0).forEach(cell
                        -> 
                        {
                            sheet.autoSizeColumn(cell.getColumnIndex(), true);
                });
            }
            */
        }
    }

    private void createEquipHdr(String hdrTitle)
    {
        //TODO HEADING TYPE 1 - for each device type *******************
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 0, 10);
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);

        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE9), row);
        row.getCell(0).setCellValue(createHelper.createRichTextString(hdrTitle));
    }

    private void createEquipHdrWithGrey(String hdrTitle, String greyHdrTitle)
    {
        //HEADING TYPE 2 **********************************
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 6, styleMap.get(STYLE9), row);
        ExcelStyleGenerator.createCell(6, 8, styleMap.get(STYLE10), row);
        ExcelStyleGenerator.createCell(8, 11, styleMap.get(STYLE9), row);
        row.getCell(0).setCellValue(createHelper.createRichTextString(hdrTitle));
        row.getCell(6).setCellValue(createHelper.createRichTextString(greyHdrTitle));
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 6, 7);
    }

    private void createSubHeading1(List<String> stringList, Row row)
    {
        int count = 0;
        for (String str : stringList)
        {
            row.getCell(count).setCellValue(createHelper.createRichTextString(str));
            count++;
        }
    }

    private void createSubHeading2(String hdrTitle, String hdrTitle2)
    {
        //HEADING TYPE 3 - Shelters and transceivers ************
        Row row = sheet.createRow((short) rowNum);
        row.setHeight((short) 300);
        ExcelStyleGenerator.createCell(0, 11, styleMap.get(STYLE8), row);

        List<String> stringList = new ArrayList<>();
        stringList.add(hdrTitle);
        stringList.add(hdrTitle2);
        stringList.add("");
        stringList.add("");
        stringList.add(STATUS_LABEL);
        stringList.add(OWNER_LABEL);
        stringList.add("");
        stringList.add(NOTES_LABEL);

        createSubHeading1(stringList, row);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 1, 3);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 5, 6);
        ExcelStyleGenerator.mergeCells(sheet, rowNum, rowNum, 7, 10);
    }
}
