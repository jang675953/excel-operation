package com.klein.easypoi.domain.one2many;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.AllArgsConstructor;
import lombok.Data;

@ExcelTarget("car")
@AllArgsConstructor
@Data
public class Car {
    @Excel(name = "汽车品牌")
    private String brand;

    @Excel(name = "价格/万元")
    private Double price;

}
