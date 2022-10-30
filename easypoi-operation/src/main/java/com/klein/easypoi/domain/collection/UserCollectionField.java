package com.klein.easypoi.domain.collection;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import com.klein.easypoi.domain.basic.User;
import lombok.Data;

import java.util.*;

@ExcelTarget("user")
@Data
public class UserCollectionField extends User {

    @Excel(name = "技能", width = 60, orderNum = "5")
    private List<String> skills;

    public static List<UserCollectionField> generateCustom() {
        List<UserCollectionField> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserCollectionField user = new UserCollectionField();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            user.setSkills(Arrays.asList("Java", "Python", "Go"));
            userList.add(user);
        }
        return userList;
    }
}
