package com.klein.easypoi.domain.one2many;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.Data;

import java.util.*;

@ExcelTarget("user")
@Data
public class UserObjectCollectionField {

    @Excel(name = "姓名", width = 30, orderNum = "1", needMerge = true)
    private String name;

    @Excel(name = "年龄", width = 10, orderNum = "2", needMerge = true)
    private Integer age;

    @Excel(name = "性别", width = 10, orderNum = "3", replace = {"女_0", "男_1"}, needMerge = true)
    private Integer sex;

    @Excel(name = "生日", width = 30, orderNum = "4", format = "yyyy-MM-dd HH:mm:ss", needMerge = true)
    private Date birthday;

    @ExcelCollection(name = "拥有的汽车", orderNum = "5")
    private List<Car> cars;

    public static List<UserObjectCollectionField> generate(){
        List<UserObjectCollectionField> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserObjectCollectionField user = new UserObjectCollectionField();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            user.setCars(Arrays.asList(new Car("宝马", 30.0), new Car("保时捷", 120.0), new Car("库里南", 580.0)));
            userList.add(user);
        }
        return userList;
    }
}
