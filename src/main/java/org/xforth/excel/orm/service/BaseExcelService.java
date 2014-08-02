package org.xforth.excel.orm.service;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xforth.excel.orm.entity.*;
import org.xforth.excel.orm.exception.SheetNotFoundException;
import org.xforth.excel.orm.util.ExcelUtils;
import org.xforth.excel.orm.util.ReflectionUtils;

import javax.annotation.PostConstruct;
import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

/***
 * Excel Service:
 * 1.excel file -> bean list
 * 2.bean list -> export excel
 * 3.export excel template
 * @param <T> extends BaseExcelEntity
 */
public class BaseExcelService<T extends BaseExcelEntity>  extends ExcelGenerator {

    private static final Logger logger = LoggerFactory.getLogger(BaseExcelService.class);
    protected static final int MAX_HEADER_SIZE = 20;
	private Class<T> entityClass;
    private HeaderMeta headerMeta;
    private String[] sheetNames;
    private final HashSet<String> sheetNameSet = new HashSet<>();

	public BaseExcelService() throws IllegalAccessException, InstantiationException {
		this.entityClass = ReflectionUtils.getSuperClassGenricType(getClass());
	}
    @PostConstruct
    public void init() throws IllegalAccessException, InstantiationException {
        T entityInstance = entityClass.newInstance();
        headerMeta = entityInstance.getHeaderMeta();
        sheetNames = entityInstance.getSheetMeta();
        for(String sheetName:sheetNames){
            sheetNameSet.add(sheetName);
        }
    }
	/***
	 * excelfile -> bean list
	 * 
	 * @param is：Excel file inputstream
	 * @return List<T>
	 * @throws IOException
     * @throws IllegalAccessException
     */
	protected List<T> generateEntity(InputStream is) throws IOException, IllegalAccessException, InstantiationException {
        try {
            List<T> beanList = new ArrayList<T>();
            POIFSFileSystem fs = new POIFSFileSystem(is);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            if (sheetNames != null && sheetNames.length > 0) {
                List<HSSFSheet> sheetList = ExcelUtils.getSheetsByNames(wb, sheetNames);
                for (HSSFSheet sheet : sheetList) {
                    int rowNum = sheet.getPhysicalNumberOfRows();
                    if (rowNum > 0) {
                        HeaderInfo headerInfo = ExcelUtils.buildHeaderInfo(sheet, headerMeta);
                        for (int i = 1; i < rowNum; i++) {
                            HSSFRow row = sheet.getRow(i);
                            int colNum = row.getPhysicalNumberOfCells();
                            T t = entityClass.newInstance();
                            t.setSheetName(sheet.getSheetName());
                            for (int j = 0; j < colNum; j++) {
                                HSSFCell cell = row.getCell(j);
                                String cellVal = ExcelUtils.getStringCellValue(cell);
                                ReflectionUtils.invokeSetterMethod(t, headerInfo.getPropertyNameByColumnIndex(j)
                                        , cellVal);
                            }
                            beanList.add(t);
                        }
                    }
                }
            }else{
                throw new SheetNotFoundException("sheet name not found");
            }
            return beanList;
        }catch (Exception e){
            logger.error("generateEntity exception:"+e);
            throw e;
        }finally {
            is.close();
        }
    }

	/***
	 * download excel template
	 * 
	 * @param os
	 * @throws IOException
	 */
    public void dlExcelTemplate(OutputStream os) throws IOException {
        try {
            HSSFWorkbook wb = new HSSFWorkbook();
            fillBlankWorkbook(wb);
            wb.write(os);
        }catch (Exception e){
            logger.error("dlExcelTemplate exception:"+e);
            throw e;
        }finally {
            os.flush();
            os.close();
        }
        /*response.setContentType( "application/octet-stream ");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String((fileName+".xls").getBytes("GBK"),"8859_1"));
        response.getOutputStream().close();*/
	}



    /***
	 * download excel
	 * @param sheetGroup
	 * @param os
	 * @throws Exception
	 */
    protected void dlExcel(SheetGroup sheetGroup,OutputStream os) throws Exception{
        try {
            HSSFWorkbook wb = new HSSFWorkbook();
            fillWorkbook(sheetGroup, wb);
            wb.write(os);
        }catch (Exception e){
            logger.error("dlExcel exception:"+e);
            throw e;
        }finally {
            os.flush();
            os.close();
        }
	}
    @Override
    protected void fillBlankWorkbook(HSSFWorkbook wb) {
        for(String sheetName:sheetNames){
            fillBlankSheet(wb, sheetName);
        }
    }
    @Override
    protected void fillWorkbook(SheetGroup sheetGroup, HSSFWorkbook wb) {
        List<String> headerTitleList = headerMeta.getHeaderTitle();
        for(String sheetName:sheetNames){
            List<BaseExcelEntity> entityList = sheetGroup.getBySheetName(sheetName);
            if(entityList==null){
                fillBlankSheet(wb, sheetName);
            }else {
                fillSheet(wb, headerTitleList, sheetName, entityList);
            }
        }
    }
    @Override
    protected void fillBlankSheet(HSSFWorkbook wb, String sheetName) {
        HSSFSheet sheet = wb.createSheet(sheetName);
        ExcelUtils.writeSheetHeader(sheet, headerMeta);
    }
    @Override
    protected void fillSheet(HSSFWorkbook wb, List<String> headerTitleList, String sheetName, List<BaseExcelEntity> entityList) {
        HSSFSheet sheet = wb.createSheet(sheetName);
        ExcelUtils.writeSheetHeader(sheet, headerMeta);
        int rowNum = 1;
        for (BaseExcelEntity entity : entityList) {
            HSSFRow row = sheet.createRow(rowNum++);
            for (int col = 0; col < headerTitleList.size(); col++) {
                Object rowObject = ReflectionUtils.invokeGetterMethod(entity,
                        headerMeta.getPropertyNameByHeaderTitle(headerTitleList.get(col)));
                HSSFCell cell = row.createCell(col);
                ExcelUtils.setDefaultCellStyle(cell.getCellStyle());
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                cell.setCellValue(rowObject==null?null:rowObject.toString());
            }
        }
    }

    protected SheetGroup createSheetGroup(){
        return new SheetGroup(sheetNameSet);
    }
}
