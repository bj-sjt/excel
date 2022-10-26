package com.itao.excel.util;


public class StringUtil {

    public static boolean isBlank(String str) {
        return str == null || "".equals(str);
    }
    public static boolean isNotBlank(String str) {
        return str != null && !"".equals(str);
    }
}
