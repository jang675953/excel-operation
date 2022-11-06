package com.klein.easypoi.export;

import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import com.klein.easypoi.domain.basic.User;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.stream.Collectors;

public class BigDataExportServer implements IExcelExportServer {
    @Override
    public List<Object> selectListForExcelExport(Object obj, int page) {
        if (((int) obj) == page) {
            return null;
        }
        return User.generate(1000).stream().map(user -> (Object) user).collect(Collectors.toList());
    }
}
