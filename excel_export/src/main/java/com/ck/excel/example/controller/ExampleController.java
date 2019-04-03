package com.ck.excel.example.controller;

import com.ck.excel.annotations.CreateExcel;
import com.ck.excel.annotations.TitleAndVal;
import com.ck.excel.example.entity.Person;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * 导出案例
 *
 * @author QianNianXiaoYao
 * @serial 2019/4/3
 */
@RestController
public class ExampleController {

    private static final List<Person> personList = new ArrayList<>();

    /*
     * 初始化十条需要导出的数据
     */
    static {
        Random random = new Random();
        for (int i = 0; i < 10; i++) {
            personList.add(Person.builder()
                    .id(i)
                    .name("testName " + i)
                    .age(random.nextInt(30))
                    .sex(random.nextInt(2) + "")
                    .birthday(new Date().toLocaleString())
                    .build());
        }
    }

    /**
     * 导出excel
     *
     * @return list
     * @author QianNianXiaoYao
     */
    @CreateExcel(fileName = "人员信息", titleAndVal = {
            @TitleAndVal(titleName = "ID", mapField = "id"),
            @TitleAndVal(titleName = "姓名", mapField = "name"),
            @TitleAndVal(titleName = "年龄", mapField = "age"),
            // 使用表达式
            @TitleAndVal(titleName = "性别", mapField = "sex", simpleExpression = "0=女|1=男"),
            @TitleAndVal(titleName = "生日", mapField = "birthday")
    })
    @GetMapping("/exportExcel")
    public List<Person> exportExcel() {
        return personList;
    }
}
