package org.xforth.excel.orm.service;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.xforth.excel.orm.entity.*;
import org.xforth.excel.orm.util.ExcelUtils;
import org.xforth.excel.orm.util.ReflectionUtils;
import javax.servlet.http.HttpServletResponse;
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
public class ExcelBaseService<T extends BaseExcelEntity> {

    protected static final int MAX_HEADER_SIZE = 20;
	private Class<T> entityClass;
    private static HeaderMeta headerMeta;
    private static String[] sheetNames;
    private static final HashSet<String> sheetNameSet = new HashSet<>();

	public ExcelBaseService() throws IllegalAccessException, InstantiationException {
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
	 * @param file：Excel file
	 * @return List<T>
	 * @throws IOException
     * @throws IllegalAccessException
     */
	@SuppressWarnings("unchecked")
	protected List<T> generateEntity(File file) throws IOException, IllegalAccessException, InstantiationException {
        List<T> beanList = new ArrayList<T>();
        FileInputStream fis = new FileInputStream(file);
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
    }

	/***
	 * download excel template
	 * 
	 * @param fileName
	 * @param response
	 * @throws IOException
	 */
    public void dlExcelTemplate(String fileName, HttpServletResponse response) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        for(String sheetName:sheetNames){
            HSSFSheet sheet = wb.createSheet(sheetName);
            ExcelUtils.writeSheetHeader(sheet, headerMeta);
        }
        wb.write(response.getOutputStream());
        response.setContentType( "application/octet-stream ");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String((fileName+".xls").getBytes("GBK"),"8859_1"));
        response.getOutputStream().close();
	}

	/***
	 * download excel
	 * @param fileName
	 * @param sheetGroup
	 * @param response
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
    protected void dlExcel(String fileName,SheetGroup sheetGroup,HttpServletResponse response) throws Exception{
        HSSFWorkbook wb = new HSSFWorkbook();
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
                    row.createCell(col).setCellValue(rowVal);
                }
            }
        }
        wb.write(response.getOutputStream());
        response.setContentType( "application/octet-stream ");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String((fileName+".xls").getBytes("GBK"),"8859_1"));
        response.getOutputStream().close();
	}
    protected SheetGroup createSheetGroup(){
        return new SheetGroup(sheetNameSet);
    }
}
