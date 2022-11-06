package com.klein.poi;

import com.klein.poi.domain.basic.SexEnum;
import com.klein.poi.domain.basic.User;
import com.klein.poi.domain.picture.UserPicture;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

class PoiOperationApplicationTests {

    private static final String BASE_PATH = "F:\\GitHub\\excel-operation\\poi-operation\\src\\test\\resources\\";
    private static final String FILE_NAME = "UserList.xlsx";

    private int[] COLUMN_WIDTHS = {10 * 256, 10 * 256, 10 * 256, 20 * 256};
    private String[] HEADS = {"姓名", "年龄", "性别", "生日"};

    private int[] PICTURE_COLUMN_WIDTHS = {10 * 256, 10 * 256, 10 * 256, 20 * 256, 60 * 60};
    private String[] PICTURE_HEADS = {"姓名", "年龄", "性别", "生日", "头像"};

    @Test
    void poiWrite() throws IOException {
        //1.创建HSSFWorkbook
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //2.创建工作簿
        HSSFSheet sheet = hssfWorkbook.createSheet();
        //3.设置列宽
        for (int i = 0; i < COLUMN_WIDTHS.length; i++) {
            sheet.setColumnWidth((short) i, (short) COLUMN_WIDTHS[i]);
        }
        //4.创建标题行
        HSSFRow titleRow = sheet.createRow(0);
        //设置行高
        titleRow.setHeightInPoints((short) 20);
//        //创建Title单元格格式并设置
        HSSFCellStyle titleHssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont hssfFont = hssfWorkbook.createFont();
        hssfFont.setFontName(HSSFFont.FONT_ARIAL);
        hssfFont.setFontHeightInPoints((short) 11);
        hssfFont.setBold(Boolean.TRUE);
        titleHssfCellStyle.setFont(hssfFont);
        titleHssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleHssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        titleHssfCellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
        titleHssfCellStyle.setHidden(Boolean.FALSE);
        titleRow.setRowStyle(titleHssfCellStyle);
        for (int i = 0; i < HEADS.length; i++) {
            final HSSFCell titleRowCell = titleRow.createCell(i);
            titleRowCell.setCellStyle(titleHssfCellStyle);
            titleRowCell.setCellType(CellType.STRING);
            titleRowCell.setCellValue(HEADS[i]);
        }
        //创建行单元格格式并设置
        HSSFCellStyle rowHssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont rowHssfFont = hssfWorkbook.createFont();
        rowHssfFont.setFontName(HSSFFont.FONT_ARIAL);
        rowHssfFont.setFontHeightInPoints((short) 11);
        rowHssfCellStyle.setFont(rowHssfFont);
        rowHssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        rowHssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        rowHssfCellStyle.setWrapText(Boolean.TRUE);
        //创建单元格格式
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont cellHssfFont = hssfWorkbook.createFont();
        cellHssfFont.setFontName(HSSFFont.FONT_ARIAL);
        cellHssfFont.setFontHeightInPoints((short) 11);
        hssfCellStyle.setFont(cellHssfFont);
        hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        hssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        hssfCellStyle.setWrapText(Boolean.TRUE);
        //日期单元格格式
        HSSFCellStyle dateCellStyle = hssfWorkbook.createCellStyle();
        cellHssfFont.setFontHeightInPoints((short) 11);
        dateCellStyle.setFont(cellHssfFont);
        dateCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dateCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        dateCellStyle.setWrapText(Boolean.TRUE);
        dateCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        final List<User> userList = User.generate();
        //5.遍历数据,创建数据行
        for (User user : userList) {
            //获取最后一行的行号
            int lastRowNum = sheet.getLastRowNum();
            HSSFRow dataRow = sheet.createRow(lastRowNum + 1);
            //设置行单元格格式
            dataRow.setRowStyle(rowHssfCellStyle);
            final HSSFCell nameCell = dataRow.createCell(0);
            nameCell.setCellStyle(hssfCellStyle);
            nameCell.setCellType(CellType.STRING);
            nameCell.setCellValue(user.getName());
            final HSSFCell ageCell = dataRow.createCell(1);
            ageCell.setCellStyle(hssfCellStyle);
            ageCell.setCellType(CellType.NUMERIC);
            ageCell.setCellValue(user.getAge());
            final HSSFCell sexCell = dataRow.createCell(2);
            sexCell.setCellStyle(hssfCellStyle);
            sexCell.setCellType(CellType.STRING);
            sexCell.setCellValue(user.getSex());
            final HSSFCell birthCell = dataRow.createCell(3);
            final short builtinFormat = HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm");
            dateCellStyle.setDataFormat(builtinFormat);
            birthCell.setCellStyle(dateCellStyle);
            birthCell.setCellType(CellType.STRING);
            birthCell.setCellValue(user.getBirthday());
//            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
//            birthCell.setCellStyle(hssfCellStyle);
//            birthCell.setCellType(CellType.STRING);
//            birthCell.setCellValue(simpleDateFormat.format(user.getBirthday()));
        }
        FileOutputStream fileOut = new FileOutputStream(BASE_PATH + "basic" + File.separator + FILE_NAME);
        hssfWorkbook.write(fileOut);
        fileOut.flush();
        hssfWorkbook.close();
        fileOut.close();
    }

    @Test
    void poiWritePicture() throws IOException {
        //1.创建HSSFWorkbook
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //2.创建工作簿
        HSSFSheet sheet = hssfWorkbook.createSheet();
        //3.设置列宽
        for (int i = 0; i < PICTURE_COLUMN_WIDTHS.length; i++) {
            sheet.setColumnWidth((short) i, (short) PICTURE_COLUMN_WIDTHS[i]);
        }
        //4.创建标题行
        HSSFRow titleRow = sheet.createRow(0);
        //设置行高
        titleRow.setHeightInPoints((short) 20);
//        //创建Title单元格格式并设置
        HSSFCellStyle titleHssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont hssfFont = hssfWorkbook.createFont();
        hssfFont.setFontName(HSSFFont.FONT_ARIAL);
        hssfFont.setFontHeightInPoints((short) 11);
        hssfFont.setBold(Boolean.TRUE);
        titleHssfCellStyle.setFont(hssfFont);
        titleHssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleHssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        titleHssfCellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
        titleRow.setRowStyle(titleHssfCellStyle);
        for (int i = 0; i < PICTURE_HEADS.length; i++) {
            final HSSFCell titleRowCell = titleRow.createCell(i);
            titleRowCell.setCellStyle(titleHssfCellStyle);
            titleRowCell.setCellType(CellType.STRING);
            titleRowCell.setCellValue(PICTURE_HEADS[i]);
        }
        //创建行单元格格式并设置
        HSSFCellStyle rowHssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont rowHssfFont = hssfWorkbook.createFont();
        rowHssfFont.setFontName(HSSFFont.FONT_ARIAL);
        rowHssfFont.setFontHeightInPoints((short) 11);
        rowHssfCellStyle.setFont(rowHssfFont);
        rowHssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        rowHssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        rowHssfCellStyle.setWrapText(Boolean.TRUE);
        //创建单元格格式
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont cellHssfFont = hssfWorkbook.createFont();
        cellHssfFont.setFontName(HSSFFont.FONT_ARIAL);
        cellHssfFont.setFontHeightInPoints((short) 11);
        hssfCellStyle.setFont(cellHssfFont);
        hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        hssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        hssfCellStyle.setWrapText(Boolean.TRUE);
        //日期单元格格式
        HSSFCellStyle dateCellStyle = hssfWorkbook.createCellStyle();
        cellHssfFont.setFontHeightInPoints((short) 11);
        dateCellStyle.setFont(cellHssfFont);
        dateCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dateCellStyle.setVerticalAlignment(VerticalAlignment.CENTER.CENTER);
        dateCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        dateCellStyle.setWrapText(Boolean.TRUE);
        //创建HSSFPatriarch
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        final List<UserPicture> userList = UserPicture.generateCustom();

        //5.遍历数据,创建数据行
        for (int i = 0; i < userList.size(); i++) {
            UserPicture user = userList.get(i);
            //获取最后一行的行号
            int lastRowNum = sheet.getLastRowNum();
            HSSFRow dataRow = sheet.createRow(lastRowNum + 1);
            dataRow.setHeightInPoints(60);
            //设置行单元格格式
            dataRow.setRowStyle(rowHssfCellStyle);
            final HSSFCell nameCell = dataRow.createCell(0);
            nameCell.setCellType(CellType.STRING);
            nameCell.setCellStyle(hssfCellStyle);
            nameCell.setCellValue(user.getName());
            final HSSFCell ageCell = dataRow.createCell(1);
            ageCell.setCellStyle(hssfCellStyle);
            ageCell.setCellType(CellType.NUMERIC);
            ageCell.setCellValue(user.getAge());
            final HSSFCell sexCell = dataRow.createCell(2);
            sexCell.setCellStyle(hssfCellStyle);
            sexCell.setCellValue(SexEnum.getNameByCode(user.getSex()));
            sexCell.setCellType(CellType.STRING);
            final HSSFCell birthCell = dataRow.createCell(3);
            final short builtinFormat = HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm");
            dateCellStyle.setDataFormat(builtinFormat);
            birthCell.setCellStyle(dateCellStyle);
            birthCell.setCellValue(user.getBirthday());
            final HSSFCell avatarCell = dataRow.createCell(4);
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 30, 40, (short) 4, i + 1, (short) 5, i + 2);
            patriarch.createPicture(anchor, hssfWorkbook.addPicture(user.getAvatar(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            avatarCell.setCellStyle(hssfCellStyle);
        }
        FileOutputStream fileOut = new FileOutputStream(BASE_PATH + "picture" + File.separator + FILE_NAME);
        hssfWorkbook.write(fileOut);
        fileOut.flush();
        hssfWorkbook.close();
        fileOut.close();
    }

    @Test
    void poiRead() throws IOException, ParseException {
        List<User> users = new ArrayList<>();
        // 1、获取文件输入流
        ClassPathResource classPathResource = new ClassPathResource("basic/UserList.xlsx");
        // 2、创建HSSFWorkbook
        HSSFWorkbook workbook = new HSSFWorkbook(classPathResource.getInputStream());
        // 3、获取Sheet
        HSSFSheet sheetAt = workbook.getSheetAt(0);
        // 4、循环读取表格数据
        for (Row row : sheetAt) {
            //不读取表头
            if (row.getRowNum() == 0) {
                continue;
            }
            //读取当前行中单元格数据，索引从0开始
            final Cell nameCell = row.getCell(0);
            String name = nameCell.getStringCellValue();
            final Cell ageCell = row.getCell(1);
            double age = ageCell.getNumericCellValue();
            final Cell sexCell = row.getCell(2);
            double sex = sexCell.getNumericCellValue();
            final Cell birthCell = row.getCell(3);
            double birthday = birthCell.getNumericCellValue();
            User user = new User();
            user.setName(name);
            user.setAge(BigDecimal.valueOf(age).intValue());
            user.setSex(BigDecimal.valueOf(sex).intValue());
            user.setBirthday(DateUtil.getJavaDate(birthday));
            users.add(user);
        }
        Assertions.assertEquals(10, users.size());
        //5、关闭流
        workbook.close();
    }

    @Test
    void poiReadPicture() throws IOException {
        List<UserPicture> users = new ArrayList<>();
        // 1、获取文件输入流
        ClassPathResource classPathResource = new ClassPathResource("picture/UserList.xlsx");
        // 2、创建HSSFWorkbook
        HSSFWorkbook workbook = new HSSFWorkbook(classPathResource.getInputStream());
        // 3、获取Sheet
        HSSFSheet sheetAt = workbook.getSheetAt(0);
        // 获取所有图片
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        // 4、循环读取表格数据
        int rowCount = 0;
        for (Row row : sheetAt) {
            //不读取表头
            if (row.getRowNum() == 0) {
                continue;
            }
            //读取当前行中单元格数据，索引从0开始
            final Cell nameCell = row.getCell(0);
            String name = nameCell.getStringCellValue();
            final Cell ageCell = row.getCell(1);
            double age = ageCell.getNumericCellValue();
            final Cell sexCell = row.getCell(2);
            String sex = sexCell.getStringCellValue();
            final Cell birthCell = row.getCell(3);
            double birthday = birthCell.getNumericCellValue();
            UserPicture user = new UserPicture();
            user.setName(name);
            user.setAge(BigDecimal.valueOf(age).intValue());
            user.setSex(SexEnum.getEnumByCode(sex));
            user.setBirthday(DateUtil.getJavaDate(birthday));
            user.setAvatar(pictures.get(rowCount).getData());
            users.add(user);
            rowCount++;
        }
        Assertions.assertEquals(10, users.size());
        Assertions.assertEquals(11670, users.get(0).getAvatar().length);
        //5、关闭流
        workbook.close();
    }


    public String getCellStringValue(HSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case STRING://字符串类型
                cellValue = cell.getStringCellValue();
                if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
                    cellValue = " ";
                break;
            case NUMERIC: //数值类型
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case FORMULA: //公式
                cell.setCellType(CellType.NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case BLANK:
                cellValue = " ";
                break;
            case BOOLEAN:
                cellValue = Boolean.toString(cell.getBooleanCellValue());
                break;
            case ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
            default:
                break;
        }
        return cellValue;
    }


}
