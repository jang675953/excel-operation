package com.klein.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.klein.easypoi.domain.basic.User;
import com.klein.easypoi.domain.collection.UserCollectionField;
import com.klein.easypoi.domain.object.UserObjectField;
import com.klein.easypoi.domain.one2many.UserObjectCollectionField;
import com.klein.easypoi.domain.picture.UserPicture;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

class EasyPoiOperationApplicationTests {

    @Test
    void basicWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), User.class, User.generate());
        FileOutputStream outputStream = new FileOutputStream("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\basic\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void pictureWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1", ExcelType.HSSF), UserPicture.class, UserPicture.generateCustom());
        FileOutputStream outputStream = new FileOutputStream("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\picture\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userObjectFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserObjectField.class, UserObjectField.generateCustom());
        FileOutputStream outputStream = new FileOutputStream("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\object\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userCollectionFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserCollectionField.class, UserCollectionField.generateCustom());
        FileOutputStream outputStream = new FileOutputStream("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\collection\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userObjectCollectionFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserObjectCollectionField.class, UserObjectCollectionField.generate());
        FileOutputStream outputStream = new FileOutputStream("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\one2many\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }


}
