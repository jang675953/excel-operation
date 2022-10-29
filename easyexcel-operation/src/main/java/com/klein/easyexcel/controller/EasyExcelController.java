package com.klein.easyexcel.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.klein.easyexcel.domain.User;
import org.springframework.util.StopWatch;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Collection;

public class EasyExcelController {
    /**
     * 文件下载
     */
    @GetMapping("/easy-excel/download")
    public void download(HttpServletResponse response) throws IOException {
        StopWatch stopWatch = new StopWatch("easy-excel-download");
        stopWatch.start();
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String fileName = "EasyExcel";
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), User.class).sheet("模板").doWrite(data());
        stopWatch.stop();
        stopWatch.prettyPrint();
    }

    private Collection<?> data() {
        return null;
    }

    /**
     * 文件上传
     */
    @PostMapping("/easy-excel/upload")
    public String upload(MultipartFile file) throws IOException {
        StopWatch stopWatch = new StopWatch("easy-excel-upload");
        stopWatch.start();
        EasyExcel.read(file.getInputStream(), User.class, new ReadListener() {
            @Override
            public void invoke(Object o, AnalysisContext analysisContext) {

            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {

            }
        }).sheet().doRead();
        stopWatch.stop();
        stopWatch.prettyPrint();
        return "success";
    }
}
