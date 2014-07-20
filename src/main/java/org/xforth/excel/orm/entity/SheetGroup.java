package org.xforth.excel.orm.entity;

import org.apache.commons.lang3.StringUtils;
import org.xforth.excel.orm.exception.SheetNotFoundException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;

public class SheetGroup {
    private HashMap<String,List<BaseExcelEntity>> sheetBeanGroup = new HashMap<>();
    private HashSet<String> sheetNameSet;

    public SheetGroup(HashSet<String> sheetNameSet) {
        this.sheetNameSet = sheetNameSet;
    }

    public <T extends BaseExcelEntity> void add(T entity){
        String sheetName = entity.getSheetName();
        if(StringUtils.isBlank(sheetName)){
            throw new SheetNotFoundException("sheet name is null");
        }
        if(!sheetNameSet.contains(sheetName)){
            throw new SheetNotFoundException("sheet name not found:"+entity.getSheetName());
        }else{
            if(sheetBeanGroup.containsKey(sheetName)){
                sheetBeanGroup.get(sheetName).add(entity);
            }else{
                List<BaseExcelEntity> list = new ArrayList<BaseExcelEntity>();
                list.add(entity);
                sheetBeanGroup.put(sheetName,list);
            }
        }
    }
    public List<BaseExcelEntity> getBySheetName(String sheetName){
        return sheetBeanGroup.get(sheetName);
    }
}
