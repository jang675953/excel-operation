package com.klein.poi.domain.basic;

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

public class User {
    private String name;

    private Integer age;

    private Integer sex;

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
