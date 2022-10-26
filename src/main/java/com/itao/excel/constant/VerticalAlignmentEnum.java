package com.itao.excel.constant;

import com.itao.excel.util.StringUtil;
import lombok.Getter;

@Getter
public enum VerticalAlignmentEnum {

    TOP("top"),
    CENTER("center"),
    BOTTOM("bottom"),
    JUSTIFY("justify"),
    DISTRIBUTED("distributed");

    private final String alignment;

    VerticalAlignmentEnum(String alignment) {
        this.alignment = alignment;
    }

    public static int getCode(String alignment){
        if (StringUtil.isBlank(alignment)) {
            return CENTER.ordinal();
        }
        VerticalAlignmentEnum[] values = VerticalAlignmentEnum.values();
        for(VerticalAlignmentEnum verticalAlignmentEnum : values) {
            if (verticalAlignmentEnum.getAlignment().equals(alignment)) {
                return verticalAlignmentEnum.ordinal();
            }
        }
        return CENTER.ordinal();
    }
}
