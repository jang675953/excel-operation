package com.klein.easyexcel.domain.annotation;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.BooleanEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
import com.alibaba.excel.enums.poi.VerticalAlignmentEnum;
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
public class User {

    @ColumnWidth(30)
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @ColumnWidth(10)
    @ExcelProperty(value = "年龄", index = 1)
    private Integer age;

    @ColumnWidth(10)
    @ExcelProperty(value = "性别", index = 2)
    private Integer sex;

    @ColumnWidth(30)
    @DateTimeFormat("yyyy-MM-dd HH:mm:ss")
    @ExcelProperty(value = "生日", index = 3)
    private Date birthday;

    public static List<User> generate() {
        return generate(10);
    }

    public static List<User> generate(int size) {
        List<User> userList = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            User user = new User();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            userList.add(user);
        }
        return userList;
    }
}
