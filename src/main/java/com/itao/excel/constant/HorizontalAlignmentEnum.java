package com.itao.excel.constant;

import com.itao.excel.util.StringUtil;
import lombok.Getter;

@Getter
public enum HorizontalAlignmentEnum {
    GENERAL("general"),
    LEFT("left"),
    CENTER("center"),
    RIGHT("right"),
    FILL("fill"),
    JUSTIFY("justify"),
    CENTER_SELECTION("center_selection"),
    DISTRIBUTED("distributed");
    private final String alignment;
    HorizontalAlignmentEnum(String alignment){
        this.alignment = alignment;
    }

    public static int getCode(String alignment){
        if (StringUtil.isBlank(alignment)) {
            return GENERAL.ordinal();
        }
        HorizontalAlignmentEnum[] values = HorizontalAlignmentEnum.values();
        for(HorizontalAlignmentEnum horizontalAlignmentEnum : values) {
            if (horizontalAlignmentEnum.getAlignment().equals(alignment)) {
                return horizontalAlignmentEnum.ordinal();
            }
        }
        return GENERAL.ordinal();
    }
}
