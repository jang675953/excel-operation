package com.klein.easyexcel.domain.converter;

import lombok.Getter;

@Getter
public enum SexEnum {
    MALE(0, "男"),
    FEMALE(1, "女");

    private final Integer code;
    private final String name;

    SexEnum(Integer code, String name) {
        this.code = code;
        this.name = name;
    }

    public static SexEnum getEnumByCode(Integer code) {
        for (SexEnum sexEnum : SexEnum.values()) {
            if (sexEnum.code.equals(code)) {
                return sexEnum;
            }
        }
        return null;
    }

    public static SexEnum getEnumByName(String name) {
        for (SexEnum sexEnum : SexEnum.values()) {
            if (sexEnum.name.equals(name)) {
                return sexEnum;
            }
        }
        return null;
    }
}
