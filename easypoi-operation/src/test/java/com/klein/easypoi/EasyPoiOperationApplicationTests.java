package com.klein.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.klein.easypoi.domain.basic.User;
import com.klein.easypoi.domain.collection.UserCollectionField;
import com.klein.easypoi.domain.object.UserObjectField;
import com.klein.easypoi.domain.one2many.UserObjectCollectionField;
import com.klein.easypoi.domain.picture.UserPicture;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

class EasyPoiOperationApplicationTests {

    private static final String BASE_PATH = "F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\";

    /**
     * 写
     */
    @Test
    void basicWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), User.class, User.generate());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "basic\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void pictureWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserPicture.class, UserPicture.generateCustom());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "picture\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userObjectFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserObjectField.class, UserObjectField.generateCustom());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "object\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userCollectionFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserCollectionField.class, UserCollectionField.generateCustom());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "collection\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void userObjectCollectionFieldWrite() throws IOException {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), UserObjectCollectionField.class, UserObjectCollectionField.generate());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "one2many\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    @Test
    void basicRead() throws Exception {
        ImportParams importParams = new ImportParams();
        importParams.setTitleRows(1);
        importParams.setHeadRows(1);
        List<User> userList = ExcelImportUtil.importExcel(new ClassPathResource("basic/userList.xls").getInputStream(), User.class, importParams);
        Assertions.assertEquals(10, userList.size());
        Assertions.assertNotNull(userList.get(0));
    }

    @Test
    void userPictureRead() throws Exception {
        ImportParams importParams = new ImportParams();
        importParams.setTitleRows(1);
        importParams.setHeadRows(1);
        List<UserPicture> userList = ExcelImportUtil.importExcel(new ClassPathResource("picture/userList.xlsx").getInputStream(), UserPicture.class, importParams);
        Assertions.assertEquals(10, userList.size());
        Assertions.assertNotNull(userList.get(0));
    }


}
