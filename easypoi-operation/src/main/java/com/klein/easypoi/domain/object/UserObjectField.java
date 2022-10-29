package com.klein.easypoi.domain.object;

import cn.afterturn.easypoi.excel.annotation.ExcelEntity;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import com.klein.easypoi.domain.basic.User;
import lombok.Data;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@Data
@ExcelTarget("user")
public class UserObjectField extends User {

    @ExcelEntity(name = "address")
    private Address address;

    public static List<UserObjectField> generateCustom(){
        List<UserObjectField> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserObjectField user = new UserObjectField();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            user.setAddress(new Address("杭州", "滨江", "阡陌路"));
            userList.add(user);
        }
        return userList;
    }
}
