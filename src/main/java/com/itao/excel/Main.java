package com.itao.excel;

import com.itao.excel.parse.Dom4jParser;
import com.itao.excel.parse.Parser;
import com.itao.excel.parse.SaxParser;
import com.itao.excel.parse.VelocityParse;

import java.io.File;
import java.io.Reader;
import java.util.HashMap;
import java.util.Map;

public class Main {

    public static void main(String[] args) {
        Map<String, Object> context = new HashMap<>();
        context.put("name1", "测试11");
        context.put("name2", "测试22");
        context.put("name3", "测试33");
        context.put("name4", "测试44");
        Reader reader = VelocityParse.parse(context,"xml/demo.xml");
        Parser parser = new Dom4jParser();
        parser.parse(reader);
        parser.write(new File("D://1.xlsx"));
    }

    private static void m1() {
        Map<String, Object> context = new HashMap<>();
        context.put("name1", "测试11");
        context.put("name2", "测试22");
        Reader reader = VelocityParse.parse(context,"xml/demo.xml");
        Parser parser = new Dom4jParser();
        parser.parse(reader);
        parser.write(new File("D://1.xlsx"));
    }
}
