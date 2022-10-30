package com.klein.easypoi.domain.basic;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@AllArgsConstructor
@NoArgsConstructor
@Data
@ExcelTarget("user")
public class User {
    @Excel(name = "姓名", width = 30, orderNum = "1")
    private String name;

    @Excel(name = "年龄", width = 10, orderNum = "2")
    private Integer age;

    @Excel(name = "性别", width = 10, orderNum = "3", replace = {"女_0", "男_1"})
    private Integer sex;

    @Excel(name = "生日", width = 30, orderNum = "4", format = "yyyy-MM-dd HH:mm:ss")
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
