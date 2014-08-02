package org.xforth.excel.orm.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.xforth.excel.orm.entity.BaseExcelEntity;
import org.xforth.excel.orm.entity.SheetGroup;

import java.util.List;

public abstract class ExcelGenerator {
    protected abstract void fillBlankWorkbook(HSSFWorkbook wb);
    protected abstract void fillWorkbook(SheetGroup sheetGroup, HSSFWorkbook wb);
    protected abstract void fillBlankSheet(HSSFWorkbook wb, String sheetName) ;
    protected abstract void fillSheet(HSSFWorkbook wb, List<String> headerTitleList, String sheetName, List<BaseExcelEntity> entityList);
}
