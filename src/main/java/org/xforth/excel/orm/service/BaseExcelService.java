package org.xforth.excel.orm.service;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.xforth.excel.orm.entity.*;
import org.xforth.excel.orm.util.ExcelUtils;
import org.xforth.excel.orm.util.ReflectionUtils;
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
public class BaseExcelService<T extends BaseExcelEntity> {

    protected static final int MAX_HEADER_SIZE = 20;
	private Class<T> entityClass;
    private static HeaderMeta headerMeta;
    private static String[] sheetNames;
    private static final HashSet<String> sheetNameSet = new HashSet<>();

	public BaseExcelService() throws IllegalAccessException, InstantiationException {
		this.entityClass = ReflectionUtils.getSuperClassGenricType(getClass());
        BaseExcelEntity entityInstance = entityClass.newInstance();
        headerMeta = entityInstance.getHeaderMeta();
        sheetNames = entityInstance.getSheetMeta();
        for(String sheetName:sheetNames){
            sheetNameSet.add(sheetName);
        }
	}

	/***
	 * excelfile -> bean list
	 * 
	 * @param FileInputStream：Excel file inputstream
	 * @return List<T>
	 * @throws IOException
     * @throws IllegalAccessException
     */
	@SuppressWarnings("unchecked")
	protected List<T> generateEntity(FileInputStream fis) throws IOException, IllegalAccessException, InstantiationException {
        try {
            List<T> beanList = new ArrayList<T>();
            POIFSFileSystem fs = new POIFSFileSystem(fis);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            if (sheetNames == null || sheetNames.length <= 0) {
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
                                String cellVal = cell.getStringCellValue();
                                ReflectionUtils.invokeSetterMethod(t, headerInfo.getPropertyNameByColumnIndex(j)
                                        , cellVal);
                            }
                            beanList.add(t);
                        }
                    }
                }
            }
            return beanList;
        }finally {
            fis.close();
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
        }finally {
            os.close();
        }
        /*response.setContentType( "application/octet-stream ");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String((fileName+".xls").getBytes("GBK"),"8859_1"));
        response.getOutputStream().close();*/
	}

    private void fillBlankWorkbook(HSSFWorkbook wb) {
        for(String sheetName:sheetNames){
            HSSFSheet sheet = wb.createSheet(sheetName);
            ExcelUtils.writeSheetHeader(sheet, headerMeta);
        }
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
        }finally {
            os.close();
        }
	}

    private void fillWorkbook(SheetGroup sheetGroup, HSSFWorkbook wb) {
        List<String> headerTitleList = headerMeta.getHeaderTitle();
        for(String sheetName:sheetNames){
            HSSFSheet sheet = wb.createSheet(sheetName);
            ExcelUtils.writeSheetHeader(sheet, headerMeta);
            List<BaseExcelEntity> entityList = sheetGroup.getBySheetName(sheetName);
            int rowNum = 1;
            for(BaseExcelEntity entity:entityList){
                HSSFRow row = sheet.createRow(rowNum++);
                for(int col=0;col<headerTitleList.size();col++) {
                    String rowVal = ReflectionUtils.invokeGetterMethod(entity,
                            headerMeta.getPropertyNameByHeaderTitle(headerTitleList.get(col))).toString();
                    HSSFCell cell = row.createCell(col);
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                    cell.setCellValue(rowVal);
                }
            }
        }
    }

    protected SheetGroup createSheetGroup(){
        return new SheetGroup(sheetNameSet);
    }
}
