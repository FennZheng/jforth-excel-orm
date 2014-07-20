package org.xforth.excel.orm.test;

import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.AbstractJUnit4SpringContextTests;
import org.xforth.excel.orm.test.service.TestExcelService;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Paths;

@ContextConfiguration(locations = { "classpath*:test.xml" })
public class TestClient extends AbstractJUnit4SpringContextTests {
    @Autowired
    private TestExcelService testExcelService;
    @Test
    public void testImportExcel() throws URISyntaxException, IllegalAccessException, IOException, InstantiationException {
        File file = Paths.get(this.getClass().getClassLoader().getResource("test-import.xls").toURI()).toFile();
        testExcelService.importExcel(file);
    }
    @Test
    public void testExportExcelTemplate() throws URISyntaxException, IllegalAccessException, IOException, InstantiationException {
        testExcelService.dlExcelTemplate(new File("test1.xls"));
    }
    @Test
    public void testExportExcel() throws Exception {
        testExcelService.dlExcel(new File("test2.xls"));
    }

}
