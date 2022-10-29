package com.klein.easypoi.domain.picture;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import com.klein.easypoi.domain.basic.User;
import lombok.Data;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@Data
@ExcelTarget("user")
public class UserPicture extends User {
    @Excel(name = "头像信息", type = 2, width = 20, height = 35)
    private String avatar;


    public static List<UserPicture> generateCustom() {
        List<UserPicture> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserPicture user = new UserPicture();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            user.setAvatar("F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\picture\\OIP-C.jpg");
            userList.add(user);
        }
        return userList;
    }
}
