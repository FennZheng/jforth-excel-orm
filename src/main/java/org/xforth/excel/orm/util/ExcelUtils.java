package org.xforth.excel.orm.util;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.xforth.excel.orm.entity.HeaderInfo;
import org.xforth.excel.orm.entity.HeaderMeta;
import org.xforth.excel.orm.exception.HeaderNotMatchException;
import org.xforth.excel.orm.exception.SheetNotFoundException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ExcelUtils {
    public static List<HSSFSheet> getSheetsByNames(HSSFWorkbook wb,String[] sheetNames){
        List<HSSFSheet> sheetList = new ArrayList<>();
        for(String sheetName:sheetNames) {
            HSSFSheet sheet = wb.getSheet(sheetName);
            if(sheet==null)
                throw new SheetNotFoundException("sheet could not found by name:"+sheetName);
            sheetList.add(sheet);
        }
        return sheetList;
    }
    public static HeaderInfo buildHeaderInfo(HSSFSheet sheet,HeaderMeta headerMeta){
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
}
