package com.klein.easypoi.export;

import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import com.klein.easypoi.domain.basic.User;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.stream.Collectors;

@Component
public class CustomExportServer implements IExcelExportServer {
    @Override
    public List<Object> selectListForExcelExport(Object o, int i) {
        return User.generate(10000).stream().map(user -> (Object) user).collect(Collectors.toList());
    }
}
