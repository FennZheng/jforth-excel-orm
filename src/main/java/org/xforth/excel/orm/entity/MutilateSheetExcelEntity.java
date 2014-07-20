package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.exception.SheetNotFoundException;

public abstract class MutilateSheetExcelEntity {
    protected String sheetName;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        if(validateSheetName(sheetName)) {
            this.sheetName = sheetName;
        }else{
            throw new SheetNotFoundException("sheetName not found :"+sheetName);
        }
    }
    protected abstract boolean validateSheetName(String sheetName);
}
