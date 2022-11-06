package com.klein.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;
import com.klein.easypoi.domain.basic.User;
import com.klein.easypoi.domain.collection.UserCollectionField;
import com.klein.easypoi.domain.object.UserObjectField;
import com.klein.easypoi.domain.one2many.UserObjectCollectionField;
import com.klein.easypoi.domain.picture.UserPicture;
import com.klein.easypoi.domain.verify.UserVerify;
import com.klein.easypoi.domain.verify.VerifyGroupOne;
import com.klein.easypoi.export.BigDataExportServer;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.StopWatch;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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
    void templateWrite() throws IOException {
        ClassPathResource classPathResource = new ClassPathResource("template/userListTemplate.xlsx");
        final TemplateExportParams templateExportParams = new TemplateExportParams(classPathResource.getPath(), "sheet1");
        final List<User> userList = User.generate();
        final List<Map<String, Object>> mapList = userList.stream().map(user -> {
            final Map<String, Object> objectMap = JSONObject.parseObject(JSON.toJSONString(user), new TypeReference<Map<String, Object>>() {
            });
            return objectMap;
        }).collect(Collectors.toList());
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("nowDate", new Date());
        dataMap.put("userList", userList);
        Workbook workbook = ExcelExportUtil.exportExcel(templateExportParams, dataMap);
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "template\\userList.xls");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }


    @Test
    public void bigDataExport() throws Exception {
        StopWatch stopWatch = new StopWatch("大数据测试");
        stopWatch.start();
        ExportParams params = new ExportParams("大数据测试", "测试");
        Workbook workbook = ExcelExportUtil.exportBigExcel(params, User.class, new BigDataExportServer(),10);
        FileOutputStream outputStream = new FileOutputStream(BASE_PATH + "bigdata\\userList.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
        stopWatch.stop();
        System.out.println(stopWatch.prettyPrint());
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
