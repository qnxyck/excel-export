package com.ck.excel.example.entity;

import lombok.Builder;
import lombok.Data;

/**
 * @author QianNianXiaoYao
 * @serial 2019/4/3
 */
@Data
@Builder
public class Person {

    private Integer id;

    // 姓名
    private String name;

    // 年龄
    private Integer age;

    // 性别
    private String sex;

    // 生日
    private String birthday;

}
