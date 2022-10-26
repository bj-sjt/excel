package com.itao.excel.parse;

import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.nio.file.Files;

/**
 * sax解析
 */
public class SaxParser implements Parser {

    private final SaxHandler saxHandler;

    public SaxParser() {
        this.saxHandler = new SaxHandler();
    }
    @Override
    public void parse(String path) {
        parse(new File(path));
    }

    @Override
    public void parse(File xmlFile) {
        try {
            SAXParserFactory spFactory = SAXParserFactory.newInstance();
            spFactory.setFeature("http://xml.org/sax/features/validation", false);
            spFactory.setValidating(false);
            SAXParser sParser = spFactory.newSAXParser();
            XMLReader xr = sParser.getXMLReader();
            xr.setErrorHandler(saxHandler);
            xr.setContentHandler(saxHandler);
            xr.parse(new InputSource(Files.newInputStream(xmlFile.toPath())));
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void parse(Reader reader) {
        try {
            SAXParserFactory spFactory = SAXParserFactory.newInstance();
            spFactory.setFeature("http://xml.org/sax/features/validation", false);
            spFactory.setValidating(false);
            SAXParser sParser = spFactory.newSAXParser();
            XMLReader xr = sParser.getXMLReader();
            xr.setErrorHandler(saxHandler);
            xr.setContentHandler(saxHandler);
            xr.parse(new InputSource(reader));
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void write(OutputStream outputStream) {
        saxHandler.write(outputStream);
    }

    @Override
    public void write(File file) {
        saxHandler.write(file);
    }
}
