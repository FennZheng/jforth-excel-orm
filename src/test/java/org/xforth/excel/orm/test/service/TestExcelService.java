package org.xforth.excel.orm.test.service;

import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xforth.excel.orm.entity.SheetGroup;
import org.xforth.excel.orm.service.BaseExcelService;
import org.xforth.excel.orm.test.entity.TestExcelEntity;

import java.io.*;
import java.util.List;

public class TestExcelService extends BaseExcelService<TestExcelEntity> {
    private static final Logger logger = LoggerFactory.getLogger(TestExcelService.class);
    public TestExcelService() throws IllegalAccessException, InstantiationException {
    }
    public void dlExcelTemplate(File file) throws IOException {
        FileOutputStream fos = new FileOutputStream(file);
        dlExcelTemplate(fos);
    }
    public SheetGroup mockSheetGroup(){
        TestExcelEntity entity = new  TestExcelEntity();
        entity.setSheetName("USA");
        entity.setFirstName("Tim");
        entity.setSecondName("Wang");
        SheetGroup sheetGroup = createSheetGroup();
        sheetGroup.add(entity);
        return sheetGroup;
    }
    public void dlExcel(File file) throws Exception {
        FileOutputStream fos = new FileOutputStream(file);
        dlExcel(mockSheetGroup(),fos);
    }
    public void importExcel(File file) throws IOException, IllegalAccessException, InstantiationException {
        FileInputStream fis = new FileInputStream(file);
        List<TestExcelEntity> resultList = generateEntity(fis);
        logger.info(JSON.toJSONString(resultList));
    }
}
