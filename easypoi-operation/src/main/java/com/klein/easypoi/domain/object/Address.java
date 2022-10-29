package com.klein.easypoi.domain.object;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.AllArgsConstructor;
import lombok.Data;

@ExcelTarget("address")
@AllArgsConstructor
@Data
public class Address {
    @Excel(name = "省", width = 30, orderNum = "5")
    private String province;

    @Excel(name = "市", width = 30, orderNum = "6")
    private String city;

    @Excel(name = "县", width = 30, orderNum = "7")
    private String county;
}
