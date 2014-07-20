package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.exception.SheetNotFoundException;

import java.io.Serializable;

public abstract class MutilateSheetExcelEntity implements Serializable{
    private static final long serialVersionUID = -2674187278930721946L;
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
