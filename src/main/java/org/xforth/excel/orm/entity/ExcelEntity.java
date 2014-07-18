package org.xforth.excel.orm.entity;

import org.xforth.excel.orm.annotation.ExcelHeader;
import org.xforth.excel.orm.annotation.ExcelSheet;
import org.xforth.orm.excel.anotation.ExcelHeader;

import java.lang.reflect.Field;
import java.util.*;

/***
 * Excel Entity
 * @author vernon.zheng
 * @version 1.0.1
 */
public class ExcelEntity implements java.io.Serializable{
    private static final long serialVersionUID = -8158635042910595602L;
    private final HashSet<String> excelHeaderTitleSet = new HashSet<>();
    public HashSet<String> getHeaderMeta(){
        Class<?> cla = this.getClass();
        Field[] fields = cla.getDeclaredFields();
        for(Field field:fields){
            if(field.isAnnotationPresent(ExcelHeader.class)){
                ExcelHeader fieldAnno = field.getAnnotation(ExcelHeader.class);
                if(excelHeaderTitleSet.contains(fieldAnno.title())){
                    throw new IllegalArgumentException("duplicate ExcelHeader:"+fieldAnno.title());
                }else{
                    excelHeaderTitleSet.add(fieldAnno.title());
                }
            }
        }
        return excelHeaderTitleSet;
    }
    public String[] getSheetMeta(){
        Class<?> cla = this.getClass();
        if(cla.isAnnotationPresent(ExcelSheet.class)){
            ExcelSheet fieldAnno = cla.getAnnotation(ExcelSheet.class);
            return fieldAnno.names();
        }
        return null;
    }
    /***
     * 得到含有Excel列头顺序，Excel列名和对应Entity属性的有序List
     * String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName
     * @return List<String[]>
     * @throws Exception
     */
    public List<String[]> toOrderList() throws Exception {
        List<String[]> resultList = new ArrayList<String[]>();
        Class<?> cla = this.getClass();
        Field[] fields = cla.getDeclaredFields();
        HashMap<Integer, String> orderHeaderMap = new HashMap<Integer, String>();
        HashMap<String, String> headerPropertyMap = new HashMap<String,String>();
        int count = 0;
        for (int i = 0; i < fields.length; i++) {
            if(fields[i].isAnnotationPresent(ExcelHeader.class)){
                ExcelHeader fieldAnnotation = fields[i]
                        .getAnnotation(ExcelHeader.class);
                if(orderHeaderMap.containsKey(fieldAnnotation.headerOrder())){
                    throw new Exception("entity annotation exception: repeat HeaderOrder "+fieldAnnotation.headerOrder());
                }else{
                    orderHeaderMap.put(fieldAnnotation.headerOrder(), fieldAnnotation.title());
                    headerPropertyMap.put(fieldAnnotation.title(), fields[i].getName());
                }
                count++;
            }
        }
        //根据HeaderOrder排序
        for(int i=0;i<count;i++){
            String[] anno = new String[3];
            anno[0] = String.valueOf(i);
            anno[1] = orderHeaderMap.get(i);
            anno[2] = headerPropertyMap.get(anno[1]);
            if(anno[1]==null){
                throw new Exception("entity annotation exception: HeaderOrder "+i+" is null");
            }
            resultList.add(anno);
        }
        return resultList;
    }
}
