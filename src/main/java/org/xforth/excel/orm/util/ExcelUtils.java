package org.xforth.excel.orm.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.xforth.excel.orm.entity.BaseExcelEntity;
import org.xforth.excel.orm.entity.HeaderInfo;
import org.xforth.excel.orm.entity.HeaderMeta;
import org.xforth.excel.orm.exception.HeaderNotMatchException;
import org.xforth.excel.orm.exception.SheetNotFoundException;

import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {
    public static List<HSSFSheet> getSheetsByNames(final HSSFWorkbook wb,final String[] sheetNames){
        List<HSSFSheet> sheetList = new ArrayList<>();
        for(String sheetName:sheetNames) {
            HSSFSheet sheet = wb.getSheet(sheetName);
            if(sheet==null)
                throw new SheetNotFoundException("sheet could not found by name:"+sheetName);
            sheetList.add(sheet);
        }
        return sheetList;
    }
    public static HeaderInfo buildHeaderInfo(final HSSFSheet sheet,final HeaderMeta headerMeta){
        HeaderInfo headerInfo = new HeaderInfo(headerMeta);
        HSSFRow headerRow = sheet.getRow(0);
        int headerCount = headerRow.getPhysicalNumberOfCells();
        for(int r=0;r<headerCount;r++){
            String headerName = headerRow.getCell(r).toString();
            if(headerMeta.contains(headerName)){
                headerInfo.put(r,headerName);
            }
        }
        if(!headerInfo.validate()){
            throw new HeaderNotMatchException("header info not match:"+headerInfo.toString());
        }else{
            return headerInfo;
        }
    }
    public static void writeSheetHeader(final HSSFSheet sheet,final HeaderMeta headerMeta){
        HSSFRow headerRow = sheet.createRow(0);
        List<String> headerTitleList = headerMeta.getHeaderTitle();
        for(int i=0;i<headerTitleList.size();i++){
            HSSFCell cell = headerRow.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(headerTitleList.get(i));
            setHeaderCellStyle(cell.getCellStyle());
            cell.getCellStyle().setLocked(true);
        }
    }
    public static void setHeaderCellStyle(HSSFCellStyle cellStyle){
        //cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        //cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
    }
    public static void setDefaultCellStyle(HSSFCellStyle cellStyle){
        //cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        //cellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
    }
    public static String getStringCellValue(HSSFCell cell) {
        String strCell = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        if (cell == null) {
            return "";
        }
        return strCell;
    }

    public static String getDateCellValue(HSSFCell cell) {
        String result = "";
        try {
            int cellType = cell.getCellType();
            if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
                Date date = cell.getDateCellValue();
                result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1)
                        + "-" + date.getDate();
            } else if (cellType == HSSFCell.CELL_TYPE_STRING) {
                String date = getStringCellValue(cell);
                result = date.replaceAll("[年月]", "-").replace("日", "").trim();
            } else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
                result = "";
            }
        } catch (Exception e) {
            System.out.println("日期格式不正确!");
            e.printStackTrace();
        }
        return result;
    }

    public static String getCellFormatValue(HSSFCell cell) {
        String cellvalue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                case HSSFCell.CELL_TYPE_FORMULA: {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        //：2011-10-12 0:00:00
                        //cellvalue = cell.getDateCellValue().toLocaleString();

                        //：2011-10-12
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);

                    }
                    // 如果是纯数字
                    else {
                        // 取得当前Cell的数值
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case HSSFCell.CELL_TYPE_STRING:
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellvalue = " ";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;

    }

}
