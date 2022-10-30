package com.klein.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.klein.easyexcel.converter.CustomDateStringConverter;
import com.klein.easyexcel.domain.annotation.User;
import com.klein.easyexcel.domain.converter.UserConverter;
import com.klein.easyexcel.domain.map.UserMap;
import com.klein.easyexcel.listener.CustomReadListener;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.ClassUtils;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
class EasyExcelApplicationTests {

    private static final String BASE_PATH = "F:\\GitHub\\excel-operation\\easyexcel-operation\\src\\test\\resources\\";

    private static final String FILE_NAME = "userList.xlsx";

    @Test
    void annotationWrite() {
        EasyExcel.write(BASE_PATH + "annotation" + File.separator + FILE_NAME, User.class).sheet("sheet1").doWrite(User.generate());
    }

    @Test
    void customConverterWrite() {
        EasyExcel.write(BASE_PATH + "converter" + File.separator + FILE_NAME, User.class).sheet("sheet1").registerConverter(new CustomDateStringConverter()).doWrite(User.generate());
    }

    @Test
    void enumConverterWrite() {
        EasyExcel.write(BASE_PATH + "converter" + File.separator + "userSexEnumList.xlsx", UserConverter.class).sheet("sheet1").doWrite(UserConverter.generate());
    }

    @Test
    void mapWrite() {
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        final WriteFont writeFont = new WriteFont();
        writeFont.setFontName("宋体");
        writeFont.setFontHeightInPoints((short) 11);
        writeFont.setBold(Boolean.TRUE);
        headWriteCellStyle.setWriteFont(writeFont);
        List<WriteCellStyle> contentWriteCellStyleList = new ArrayList<>();
        WriteCellStyle contentWriteCellStyle = null;
        for (int i = 0; i < 5; i++) {
            contentWriteCellStyle = new WriteCellStyle();
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
            contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            final WriteFont contentWriteFont = new WriteFont();
            contentWriteFont.setFontName("宋体");
            contentWriteFont.setFontHeightInPoints((short) 11);
            contentWriteCellStyle.setWriteFont(contentWriteFont);
            contentWriteCellStyleList.add(contentWriteCellStyle);
        }
        EasyExcel.write(BASE_PATH + "map" + File.separator + "userMap.xlsx")
                .useDefaultStyle(Boolean.FALSE)
                .password("123456")
                .autoCloseStream(Boolean.TRUE)
                .inMemory(Boolean.TRUE)
                .sheet("sheet1")
                .registerWriteHandler(new SimpleColumnWidthStyleStrategy(20))
                .registerWriteHandler(new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyleList))
                .head(UserMap.generateHead())
                .doWrite(UserMap.generate());
    }

    @Test
    void templateWrite() throws IOException {
        EasyExcel.write(BASE_PATH + "template" + File.separator + "userMapWithTemplate.xlsx")
                .withTemplate(new ClassPathResource("template/userMapTemplate.xlsx").getInputStream())
                .autoCloseStream(Boolean.TRUE)
                .needHead(Boolean.FALSE)
                .sheet("sheet1")
                .doWrite(UserMap.generate());
    }


    @Test
    void annotationRead() {
        final CustomReadListener readListener = new CustomReadListener();
        EasyExcel.read(BASE_PATH + "annotation" + File.separator + FILE_NAME,
                User.class,
                readListener).sheet().doRead();
        Assertions.assertEquals(10, readListener.getCount());
    }

    @Test
    void synchronousRead() {
        final CustomReadListener readListener = new CustomReadListener();
        final List<User> userList = EasyExcel.read(BASE_PATH + "annotation" + File.separator + FILE_NAME,
                User.class,
                readListener).sheet().doReadSync();
        Assertions.assertEquals(10, userList.size());
    }

    @Test
    void annotationMultiRead() {
        final CustomReadListener readListener = new CustomReadListener();
        final String pathName = BASE_PATH + "multisheet" + File.separator + FILE_NAME;

        try (ExcelReader excelReader = EasyExcel.read(pathName).build()) {
            // 这里为了简单 所以注册了 同样的head 和Listener 自己使用功能必须不同的Listener
            ReadSheet readSheet1 =
                    EasyExcel.readSheet(0).head(User.class).registerReadListener(readListener).build();
            ReadSheet readSheet2 =
                    EasyExcel.readSheet(1).head(User.class).registerReadListener(readListener).build();
            // 这里注意 一定要把sheet1 sheet2 一起传进去，不然有个问题就是03版的excel 会读取多次，浪费性能
            excelReader.read(readSheet1, readSheet2);

        }
        Assertions.assertEquals(20, readListener.getCount());

    }

    @Test
    void annotationPageRead() {
        EasyExcel.read(BASE_PATH + "annotation" + File.separator + FILE_NAME, User.class, new PageReadListener<User>(dataList -> {
            Assertions.assertEquals(10, dataList.size());
            Assertions.assertTrue(ClassUtils.isAssignable(User.class, dataList.get(0).getClass()));
        })).sheet().doRead();
    }

}
