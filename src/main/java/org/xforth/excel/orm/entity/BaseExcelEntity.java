package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.annotation.ExcelHeader;
import org.xforth.excel.orm.annotation.ExcelSheet;

import java.lang.reflect.Field;
import java.util.HashSet;

public class BaseExcelEntity extends MutilateSheetExcelEntity{
    private final HeaderMeta headerMeta = new HeaderMeta();
    private final HashSet<String> sheetNameSet = new HashSet<>();
    private String[] sheetMeta = null;
    public BaseExcelEntity(){
        gernateHeaderMeta();
        generateSheetMeta();
    }
    private void gernateHeaderMeta(){
        Class<?> cla = this.getClass();
        Field[] fields = cla.getDeclaredFields();
        for(Field field:fields){
            if(field.isAnnotationPresent(ExcelHeader.class)){
                ExcelHeader fieldAnno = field.getAnnotation(ExcelHeader.class);
                if(headerMeta.contains(fieldAnno.title())){
                    throw new IllegalArgumentException("duplicate ExcelHeader:"+fieldAnno.title());
                }else{
                    headerMeta.put(fieldAnno.title(),field.getName());
                }
            }
        }
    }
    private void generateSheetMeta(){
        Class<?> cla = this.getClass();
        if(cla.isAnnotationPresent(ExcelSheet.class)){
            ExcelSheet fieldAnno = cla.getAnnotation(ExcelSheet.class);
            sheetMeta = fieldAnno.names();
            for(String sheetName:sheetMeta){
                sheetNameSet.add(sheetName);
            }
        }
    }
    public HeaderMeta getHeaderMeta(){
        return headerMeta;
    }
    public String[] getSheetMeta(){
        return sheetMeta;
    }

    @Override
    protected boolean validateSheetName(String sheetName) {
        return sheetNameSet.contains(sheetName);
    }
}
