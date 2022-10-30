package com.klein.easypoi.domain.verify;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;

/**
 * 自定义校验处理器
 */
public class CustomVerifyHandler implements IExcelVerifyHandler<UserVerify> {
    @Override
    public ExcelVerifyHandlerResult verifyHandler(UserVerify userVerify) {
        return null;
    }
}
