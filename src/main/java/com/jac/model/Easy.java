package com.jac.model;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import com.jac.service.BaiShiExcel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.NonNull;
import org.springframework.beans.factory.annotation.Required;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * @author ：jacketzc
 * Created in 2021/1/17 22:43
 */
@Data
@AllArgsConstructor
public class Easy {
    @ExcelIgnore
    private Integer id;


    @ExcelProperty("姓名")
    private String name;


    public static void main(String[] args) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException, ClassNotFoundException {
        String s = "1231231231232434234gdsgfdshfghgfbncv";

        while (true) {
            s += s;
            System.out.println(s);
        }
    }

    public void test1() {

    }

    public void test2() {

    }
}
