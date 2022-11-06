package com.klein.poi.domain.picture;

import com.klein.poi.domain.basic.User;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.util.IOUtils;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class UserPicture extends User implements Serializable {

    private byte[] avatar;


    public static List<UserPicture> generateCustom() throws IOException {
        ClassPathResource classPathResource = new ClassPathResource("picture/OIP-C.jpg");
        final InputStream inputStream = classPathResource.getInputStream();
        byte[] bytes = IOUtils.toByteArray(inputStream);
        List<UserPicture> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserPicture user = new UserPicture();
            user.setName("User" + i);
            user.setAge(18 + i);
            user.setBirthday(new Date());
            user.setSex(new Random().nextInt(2));
            user.setAvatar(bytes);
            userList.add(user);
        }
        return userList;
    }
}
