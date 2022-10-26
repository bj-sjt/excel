package com.itao.excel.bean;

import lombok.Data;

/**
 * 样式属性
 */
@Data
public class StyleAttr {

    // 样式id
    private String id;
    // 是否有边框
    private boolean border;
    // 是否换行
    private boolean wrapText;
    // 水平对齐方式
    private String hAlignment;
    // 垂直对齐方法
    private String vAlignment;
    /**
     * 前景色
     * @see com.itao.excel.constant.ColorEnum
     */
    private String foregroundColor;
    // 字体
    private FontAttr font;
}
