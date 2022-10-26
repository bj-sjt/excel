package com.itao.excel.parse;


import java.io.File;
import java.io.OutputStream;
import java.io.Reader;

public interface Parser {

    void parse(String path);

    void parse(File xmlFile);

    void parse(Reader reader);

    void write(OutputStream outputStream);

    void write(File file);
}
