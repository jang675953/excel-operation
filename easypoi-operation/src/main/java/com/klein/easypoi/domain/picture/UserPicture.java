package com.klein.easypoi.domain.picture;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import com.klein.easypoi.domain.basic.User;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@ExcelTarget("user")
@NoArgsConstructor
@AllArgsConstructor
@Data
public class UserPicture extends User implements Serializable {

    @Excel(name = "头像信息", type = 2, width = 20, height = 35, savePath = "F:\\GitHub\\excel-operation\\easypoi-operation\\src\\test\\resources\\picture")
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
