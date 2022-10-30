package com.klein.easyexcel.domain.converter;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.BooleanEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
import com.alibaba.excel.enums.poi.VerticalAlignmentEnum;
import com.klein.easyexcel.converter.SexEnumConverter;
import lombok.Data;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@ContentRowHeight(30)
@ContentFontStyle(fontName = "宋体", fontHeightInPoints = 11)
@HeadFontStyle(fontName = "宋体", fontHeightInPoints = 11, bold = BooleanEnum.TRUE)
@HeadStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, verticalAlignment = VerticalAlignmentEnum.CENTER)
@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, verticalAlignment = VerticalAlignmentEnum.CENTER)
@Data
public class UserConverter {

    @ColumnWidth(30)
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @ColumnWidth(10)
    @ExcelProperty(value = "年龄", index = 1)
    private Integer age;

    @ColumnWidth(10)
    @ExcelProperty(value = "性别", index = 2, converter = SexEnumConverter.class)
    private SexEnum sex;

    @ColumnWidth(30)
    @ExcelProperty(value = "生日", index = 3)
    private Date birthday;

    public static List<UserConverter> generate() {
        List<UserConverter> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserConverter user = new UserConverter();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(SexEnum.values()[new Random().nextInt(1)]);
            userList.add(user);
        }
        return userList;
    }
}
