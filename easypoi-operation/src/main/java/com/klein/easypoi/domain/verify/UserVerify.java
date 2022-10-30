package com.klein.easypoi.domain.verify;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import javax.validation.constraints.Max;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@Data
public class UserVerify {
    @NotNull(groups = {VerifyGroupOne.class})
    @Excel(name = "姓名")
    private String name;

    @Max(value = 28, message = "max 最大值不能超过28", groups = {VerifyGroupOne.class})
    @Excel(name = "年龄")
    private Integer age;

    @NotNull(groups = {VerifyGroupOne.class})
    @Excel(name = "性别", replace = {"女_0", "男_1"})
    private Integer sex;

    @NotNull(groups = {VerifyGroupOne.class})
    @Excel(name = "生日", format = "yyyy-MM-dd HH:mm:ss")
    private Date birthday;

    public static List<UserVerify> generate() {
        List<UserVerify> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserVerify user = new UserVerify();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            userList.add(user);
        }
        return userList;
    }
}
