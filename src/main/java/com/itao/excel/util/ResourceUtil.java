package com.itao.excel.util;

import java.net.URL;

public class ResourceUtil {

    /**
     * 获取 classpath 路径
     * @param clazz 指定类
     */
    public static String classPath(Class<?> clazz){
        URL url = clazz.getResource("/");
        if (url != null) {
            return url.getPath();
        }
        throw new NullPointerException();
    }

    /**
     * 获取指定类的包路径
     * @param clazz 指定类
     */
    public static String packagePath(Class<?> clazz){
        URL url = clazz.getResource("");
        if (url != null) {
            return url.getPath();
        }
        throw new NullPointerException();
    }


    public static String path(String name){
        URL url = ResourceUtil.class.getClassLoader().getResource(name);
        if (url == null) {
            return null;
        }
        return url.getPath();
    }
}
