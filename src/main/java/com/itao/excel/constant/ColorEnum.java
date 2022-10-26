package com.itao.excel.constant;

import com.itao.excel.util.StringUtil;
import lombok.Getter;
import lombok.ToString;

@Getter
@ToString
public enum ColorEnum {

    BLACK("black", 8),
    WHITE("white", 9),
    RED("red", 10),
    BRIGHT_GREEN("bright_green", 11),
    BLUE("blue", 12),
    YELLOW("yellow", 13),
    PINK("pink", 14),
    TURQUOISE("turquoise", 15),
    DARK_RED("dark_red", 16),
    GREEN("green", 17),
    DARK_BLUE("dark_blue", 18),
    DARK_YELLOW("dark_yellow", 19),
    VIOLET("violet", 20),
    TEAL("teal", 21),
    CORNFLOWER_BLUE("cornflower_blue", 24),
    MAROON("maroon", 25),
    LEMON_CHIFFON("lemon_chiffon", 26),
    ORCHID("orchid", 28),
    CORAL("coral", 29),
    ROYAL_BLUE("royal_blue", 30),
    LIGHT_CORNFLOWER_BLUE("light_cornflower_blue", 31),
    SKY_BLUE("sky_blue", 40),
    LIGHT_TURQUOISE("light_turquoise", 41),
    LIGHT_GREEN("light_green", 42),
    LIGHT_YELLOW("light_yellow", 43),
    PALE_BLUE("pale_blue", 44),
    ROSE("rose", 45),
    LAVENDER("lavender", 46),
    TAN("tan", 47),
    LIGHT_BLUE("light_blue", 48),
    AQUA("aqua", 49),
    LIME("lime", 50),
    GOLD("gold", 51),
    LIGHT_ORANGE("light_orange", 52),
    ORANGE("orange", 53),
    BLUE_GREY("blue_grey", 54),
    DARK_TEAL("dark_teal", 56),
    SEA_GREEN("sea_green", 57),
    DARK_GREEN("dark_green", 58),
    OLIVE_GREEN("olive_green", 59),
    BROWN("brown", 60),
    PLUM("plum", 61),
    INDIGO("indigo", 62),
    AUTOMATIC("automatic", 64);

    private final String color;
    private final int code;

    ColorEnum(String color, int code) {
        this.code = code;
        this.color = color;
    }

    public static short getColor(String color) {
        if (StringUtil.isBlank(color)) {
            return (short) BLACK.getCode();
        }
        ColorEnum[] values = ColorEnum.values();
        for (ColorEnum colorEnum : values) {
            if (colorEnum.getColor().equals(color)) {
                return (short) colorEnum.getCode();
            }
        }
        return (short) BLACK.getCode();
    }
}
