package com.itao.excel.bean;

import lombok.Data;

/**
 * 字体属性
 */
@Data
public class FontAttr {
    // 字体
    private String fontName;
    /**
     * 字体颜色
     * @see com.itao.excel.constant.ColorEnum
     */
    private String color;
    // 字号
    private short size;
    // 是否斜体
    private boolean italic;
    // 是否加粗
    private boolean bold;
    // 是否删除线
    private boolean strikeout;
}
