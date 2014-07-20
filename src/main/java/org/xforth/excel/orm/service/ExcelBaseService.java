package org.xforth.excel.orm.service;

import com.alibaba.fastjson.JSON;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.xforth.excel.orm.entity.BaseExcelEntity;
import org.xforth.excel.orm.entity.HeaderInfo;
import org.xforth.excel.orm.entity.HeaderMeta;
import org.xforth.excel.orm.entity.MutilateSheetExcelEntity;
import org.xforth.excel.orm.util.ExcelUtils;
import org.xforth.excel.orm.util.ReflectionUtils;
import org.xforth.orm.excel.anotation.ExcelSheet;
import org.xforth.orm.excel.entity.ExcelEntity;
import org.xforth.orm.excel.util.ReflectionUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
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
	protected Class<T> entityClass;
    protected HeaderMeta headerMeta;
    protected String[] sheetNames;

	public ExcelBaseService() throws IllegalAccessException, InstantiationException {
		this.entityClass = ReflectionUtils.getSuperClassGenricType(getClass());
        BaseExcelEntity entityInstance = entityClass.newInstance();
        headerMeta = entityInstance.getHeaderMeta();
        sheetNames = entityInstance.getSheetMeta();
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
        FileInputStream fis = new FileInputStream(file); // 根据excel文件路径创建文件流
        POIFSFileSystem fs = new POIFSFileSystem(fis); // 利用poi读取excel文件流
        HSSFWorkbook wb = new HSSFWorkbook(fs); // 读取excel工作簿
        if(sheetNames==null||sheetNames.length<=0) {
            List<HSSFSheet> sheetList = ExcelUtils.getSheetsByNames(wb, sheetNames);
            for(HSSFSheet sheet:sheetList){
                int rowNum = sheet.getPhysicalNumberOfRows();
                if(rowNum>0){
                    HeaderInfo headerInfo = ExcelUtils.buildHeaderInfo(sheet,headerMeta);
                    for(int i=1;i<rowNum;i++){
                        HSSFRow row = sheet.getRow(i);
                        int colNum = row.getPhysicalNumberOfCells();
                        T t = entityClass.newInstance();
                        t.setSheetName(sheet.getSheetName());
                        for(int j=0;j<colNum;j++){
                            HSSFCell cell = row.getCell(j);
                            String cellVal = cell.getStringCellValue();
                            ReflectionUtils.invokeSetterMethod(t,headerInfo.getPropertyNameByColumnIndex(j)
                            ,cellVal);
                        }
                        beanList.add(t);
                    }
                }
            }
        }
		return beanList;
	}
    /***
     * 根据Excel 生成Entity的List
     *
     * @param file：Excel 2003
     * @return List<T>
     * @throws Exception
     */
    @SuppressWarnings("unchecked")
    protected List<T> excelToList(InputStream inputStream) throws Exception {
        List<T> beanList = new ArrayList<T>();
        /** String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName**/
        List<String[]> headerOrderlist = (List<String[]>) ReflectionUtils
                .invokeMethod(entityClass.newInstance(), "toOrderList", null, null);
        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(0);
            int row = sheet.getRows();
            int col = sheet.getColumns()>headerOrderlist.size()?headerOrderlist.size():sheet.getColumns();
            for (int r = 0; r < row; r++) {
                String[] rowValue = new String[col];
                if (r == 0) {
                    for (int c = 0; c < col; c++) {
                        rowValue[c] = sheet.getCell(c, r).getContents() != null ? sheet
                                .getCell(c, r).getContents()
                                : "";
                    }
                    if (checkHeader(rowValue,headerOrderlist) == false){
                        throw new Exception(
                                "excelToList exception: excel的标题不一致，请检查");
                    }
                } else {
                    T t = entityClass.newInstance();
                    for (int c = 0; c < col; c++) {
                        rowValue[c] = sheet.getCell(c, r).getContents() != null ? sheet
                                .getCell(c, r).getContents()
                                : "";
                        ReflectionUtils.invokeSetterMethod(t,
                                headerOrderlist.get(c)[2], rowValue[c]);
                    }
                    beanList.add(t);
                }
            }
        } catch (Exception e) {
            throw new Exception("excelToList exception: " + e);
        }
        return beanList;
    }
	/***
	 * 检查Excel的列头是否与entity中的相同顺序（HeaderOrder）的title一致
	 * 
	 * @param rowValue
	 * @param headerOrderlist
	 * @return true : 一致    flase ：不一致
	 */
	protected boolean checkHeader(String[] rowValue,List<String[]> headerOrderlist) {
		try {
			if (headerOrderlist == null) {
				return false;
			}
			for (int i = 0; i < headerOrderlist.size(); i++) {
				/** String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName**/
				if (!headerOrderlist.get(i)[1].equals(rowValue[i])) {
					return false;
				}
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}
	/***
	 * 下载模板Excel
	 * 
	 * @param title
	 * @param response
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
    public void getExcelTemplate(String title,HttpServletResponse response) throws Exception {
		/** String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName**/
		List<String[]> headerOrderlist = (List<String[]>) ReflectionUtils
		.invokeMethod(entityClass.newInstance(), "toOrderList", null, null);
		WritableWorkbook wwb = null;
		ServletOutputStream os = null;
		try{
			os = response.getOutputStream();
			response.setContentType( "application/octet-stream "); 
			response.setHeader("Content-Disposition", "attachment; filename=" + new String((title+".xls").getBytes("GBK"),"8859_1"));
			wwb = Workbook.createWorkbook(os);
			WritableSheet sheet=wwb.createSheet("sheet", 0);
			
			for(int i=0;i<headerOrderlist.size();i++){
				Label l1 = new Label(i, 0, headerOrderlist.get(i)[1]);
				sheet.addCell(l1);
			}
			wwb.write();
		}catch(Exception e){
			throw new Exception("getExcelTemplate exception: "+e);
		}finally{
			if(wwb!=null){
				try {
					wwb.close();
				} catch (Exception e) {
					throw new Exception("getExcelTemplate exception: "+e);
				} 
			}
			if(os!=null){
				try {
					os.close();
				} catch (IOException e) {
					throw new Exception("getExcelTemplate exception: "+e);
				}
			}
		}
	}
	/***
	 * 根据Excel 生成Entity 的JSON
	 * 
	 * @param file Excel 2003
	 * @return for example :[{property1:value1,property2:value2},{property3:value3,property4:value4}]
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
    protected String readReturnJson(File file) throws Exception{
		List<HashMap> beanMapList = new ArrayList<HashMap>();
		String resultJson = "";
		/** String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName**/
		List<String[]> headerOrderlist = (List<String[]>) ReflectionUtils
				.invokeMethod(entityClass.newInstance(), "toOrderList", null, null);
			Workbook workbook = null;
			try {
				workbook = Workbook.getWorkbook(file);
				Sheet sheet = workbook.getSheet(0);
				int row = sheet.getRows();
				int col = sheet.getColumns()>headerOrderlist.size()?headerOrderlist.size():sheet.getColumns();
				for (int r = 0; r < row; r++) {
					String[] rowValue = new String[col];
					if (r == 0) {
						for (int c = 0; c < col; c++) {
							rowValue[c] = sheet.getCell(c, r).getContents() != null ? sheet
									.getCell(c, r).getContents()
									: "";
						}
						if (checkHeader(rowValue,headerOrderlist) == false){
							throw new Exception(
									"列头不匹配，请检查列头(Excel Header unMatched Exception)");
						}
					} else {
						HashMap<String,String> entityMap = new HashMap<String,String>();
						for (int c = 0; c < col; c++) {
							rowValue[c] = sheet.getCell(c, r).getContents() != null ? sheet
									.getCell(c, r).getContents()
									: "";
							entityMap.put(headerOrderlist.get(c)[2], rowValue[c]);
						}
						beanMapList.add(entityMap);
					}
				}
                resultJson = JSON.toJSONString(beanMapList);
			} catch (Exception e) {
				throw new Exception("excelToList exception: " + e);
			}
		return resultJson;
	}
	/***
	 * 导出Excel
	 * @param fileName
	 * @param response
	 * @param beanList
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
    protected void exportExcel(String fileName,HttpServletResponse response,List<T> beanList) throws Exception{
		/** String[0]:HeaderOrder ；String[1]: HeaderName ; String[2]: PropertyName**/
		List<String[]> headerOrderlist = (List<String[]>) ReflectionUtils
		.invokeMethod(entityClass.newInstance(), "toOrderList", null, null);
		HashMap<Integer, String> orderHeaderMap = new HashMap<Integer, String>();
		HashMap<String, String> headerPropertyMap = new HashMap<String,String>();
		for(int i=0;i<headerOrderlist.size();i++){
			String[] values = headerOrderlist.get(i);
			orderHeaderMap.put(Integer.parseInt(values[0]), values[1]);
			headerPropertyMap.put(values[1],values[2]);
		}
		ServletOutputStream os = null;
		WritableWorkbook wwb = null;
		try{
			os = response.getOutputStream();
			response.setContentType( "application/octet-stream "); 
			response.setHeader("Content-Disposition", "attachment; filename=" + new String((fileName+".xls").getBytes("GBK"),"8859_1"));
			wwb = Workbook.createWorkbook(os);
			WritableSheet sheet=wwb.createSheet("sheet", 0);
			
			WritableCell titleCell=sheet.getWritableCell(1, 1);
			for(int i=0;i<headerOrderlist.size();i++){
				Label l1 = new Label(i, 0, orderHeaderMap.get(i));
				sheet.addCell(l1);
				for(int j=0;j<beanList.size();j++){
					//如果获取不到列则为空
					String propertyName = headerPropertyMap.get(orderHeaderMap.get(i));
					Object content = ReflectionUtils.getFieldValue(beanList.get(j), propertyName);
					Label l2=new Label(i,j+1, content==null?"":content.toString());
					sheet.addCell(l2);
				}
			}
			wwb.write();
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			if(wwb!=null){
				try {
					wwb.close();
				} catch (WriteException e) {
					throw e;
				} catch (IOException e) {
					throw e;
				}
			}
			if(os!=null){
				try {
					os.close();
				} catch (IOException e) {
					throw e;
				}
			}
		}
		
	}
}
