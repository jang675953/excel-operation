package com.klein.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.klein.easypoi.domain.basic.User;
import com.klein.easypoi.domain.collection.UserCollectionField;
import com.klein.easypoi.domain.object.UserObjectField;
import com.klein.easypoi.domain.one2many.UserObjectCollectionField;
import com.klein.easypoi.domain.picture.UserPicture;
import com.klein.easypoi.domain.verify.UserVerify;
import com.klein.easypoi.domain.verify.VerifyGroupOne;
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
//        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1"), User.class, User.generate());
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("userList", "sheet1", ExcelType.HSSF), User.class, User.generate());
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "basic\\userList.xlsx");
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
    void verifyRead() throws Exception {
        ImportParams importParams = new ImportParams();
        importParams.setTitleRows(1);
        importParams.setHeadRows(1);
        importParams.setNeedVerify(true);
        importParams.setVerifyGroup(new Class[]{VerifyGroupOne.class});
        ExcelImportResult<UserVerify> excelImportResult = ExcelImportUtil.importExcelMore(new ClassPathResource("verify/userList.xls").getInputStream(), UserVerify.class, importParams);
        Assertions.assertNotNull(excelImportResult);
        final Workbook workbook = excelImportResult.getWorkbook();
        final Workbook failWorkbook = excelImportResult.getFailWorkbook();
        FileOutputStream fos = new FileOutputStream(BASE_PATH + "verify/userList_success.xlsx");
        workbook.write(fos);
        fos.close();
        FileOutputStream failFos = new FileOutputStream(BASE_PATH + "verify/userList_failed.xlsx");
        failWorkbook.write(failFos);
        fos.close();
        Assertions.assertEquals(5, excelImportResult.getList().size());
        Assertions.assertEquals(5, excelImportResult.getFailList().size());
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
