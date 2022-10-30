package com.klein.easyexcel.domain.map;

import lombok.Data;

import java.util.*;
import java.util.stream.Collectors;

@Data
public class UserMap {

    private String name;

    private Integer age;


    private Integer sex;

    private Date birthday;

    public static List<UserMap> generate() {
        List<UserMap> userMapList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserMap userMap = new UserMap();
            userMap.setName("User" + i);
            userMap.setAge(18 + i);
            userMap.setBirthday(new Date());
            userMap.setSex(new Random().nextInt(2));
            userMapList.add(userMap);
        }
        return userMapList;
    }

    public static List<List<String>> generateHead() {
        final String[] headArray = {"姓名", "年龄", "性别", "生日"};
        return Arrays.stream(headArray).map(head -> Arrays.asList(head)).collect(Collectors.toList());
    }
}
