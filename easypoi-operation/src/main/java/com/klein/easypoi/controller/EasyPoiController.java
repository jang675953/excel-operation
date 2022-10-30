package com.klein.easypoi.controller;

import cn.afterturn.easypoi.entity.vo.BigExcelConstants;
import cn.afterturn.easypoi.entity.vo.MapExcelConstants;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import cn.afterturn.easypoi.view.PoiBaseView;
import com.klein.easypoi.domain.basic.User;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.util.StopWatch;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

@Slf4j
@Controller
public class EasyPoiController {


    private IExcelExportServer excelExportServer;

    public EasyPoiController(IExcelExportServer excelExportServer) {
        this.excelExportServer = excelExportServer;
    }

    /**
     * 文件下载
     */
    @GetMapping("/easy-poi/download")
    public void download(ModelMap modelMap, HttpServletRequest request,
                         HttpServletResponse response) {
        List<ExcelExportEntity> entityList = new ArrayList<ExcelExportEntity>();
        ExcelExportEntity excelExportNameEntity = new ExcelExportEntity("姓名", "name");
        ExcelExportEntity excelExportAgeEntity = new ExcelExportEntity("年龄", "age");
        ExcelExportEntity excelExportSexEntity = new ExcelExportEntity("性别", "sex");
        ExcelExportEntity excelExportBirthdayEntity = new ExcelExportEntity("生日", "birthday");
        entityList.add(excelExportNameEntity);
        entityList.add(excelExportAgeEntity);
        entityList.add(excelExportSexEntity);
        entityList.add(excelExportBirthdayEntity);
        final ExportParams exportParams = new ExportParams();
        exportParams.setSheetName("Sheet1");
        exportParams.setType(ExcelType.HSSF);
        modelMap.put(MapExcelConstants.ENTITY_LIST, entityList);
        modelMap.put(MapExcelConstants.MAP_LIST, smallData());
        modelMap.put(MapExcelConstants.PARAMS, exportParams);
        modelMap.put(MapExcelConstants.FILE_NAME, "UserList");
        PoiBaseView.render(modelMap, request, response, MapExcelConstants.EASYPOI_MAP_EXCEL_VIEW);
    }

    @GetMapping("/easy-poi/download2")
    public String download2(ModelMap modelMap, HttpServletRequest request,
                            HttpServletResponse response) {
        final ExportParams exportParams = new ExportParams();
        exportParams.setSheetName("Sheet1");
        exportParams.setType(ExcelType.HSSF);
        modelMap.put(BigExcelConstants.CLASS, User.class);
        modelMap.put(BigExcelConstants.PARAMS, exportParams);
        //就是我们的查询参数,会带到接口中,供接口查询使用
        modelMap.put(BigExcelConstants.DATA_PARAMS, new HashMap<String, String>());
        modelMap.put(BigExcelConstants.DATA_INTER, excelExportServer);
        return BigExcelConstants.EASYPOI_BIG_EXCEL_VIEW;
    }


    private List<User> data() {
        return User.generate(10000);
    }
    private List<User> smallData() {
        return User.generate(5000);
    }

    /**
     * 文件上传
     */
    @PostMapping("/easy-poi/upload")
    public String upload(MultipartFile file) throws Exception {
        StopWatch stopWatch = new StopWatch("easy-poi-upload");
        stopWatch.start();
        ImportParams importParams = new ImportParams();
        importParams.setHeadRows(1);
        List<User> userList = ExcelImportUtil.importExcel(file.getInputStream(), User.class, importParams);
        log.debug("upload data size: {}", userList.size());
        stopWatch.stop();
        log.debug(stopWatch.prettyPrint());
        return "success";
    }
}
