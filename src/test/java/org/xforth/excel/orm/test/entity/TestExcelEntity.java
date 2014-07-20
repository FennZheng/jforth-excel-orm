package org.xforth.excel.orm.test.entity;

import org.xforth.excel.orm.annotation.ExcelHeader;
import org.xforth.excel.orm.annotation.ExcelSheet;
import org.xforth.excel.orm.entity.BaseExcelEntity;
@ExcelSheet(names={"China","USA"})
public class TestExcelEntity extends BaseExcelEntity{

    private static final long serialVersionUID = 2733593255752240901L;
    @ExcelHeader(title = "First Name")
    private String firstName;
    @ExcelHeader(title = "Second Name")
    private String secondName;
    @ExcelHeader(title = "Age")
    private String age;

    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public String getSecondName() {
        return secondName;
    }

    public void setSecondName(String secondName) {
        this.secondName = secondName;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }
}
