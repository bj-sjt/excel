package com.itao.excel.parse;

import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.RuntimeConstants;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;

import java.io.Reader;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.Map;
import java.util.Properties;

/**
 * velocity 模板解析
 */
public class VelocityParse {

    private static final Properties p = new Properties();
    static {
        // 设置输入输出编码类型。
        p.setProperty(Velocity.INPUT_ENCODING, "UTF-8");
        p.setProperty(Velocity.OUTPUT_ENCODING, "UTF-8");
        // 设置velocity资源加载方式为类路径
        p.setProperty(RuntimeConstants.RESOURCE_LOADER, "class");
        // 设置velocity资源加载方式为类路径时的处理类
        p.setProperty("class.resource.loader.class", ClasspathResourceLoader.class.getName());
        // p.setProperty("eventhandler.referenceinsertion.class", "org.apache.velocity.app.event.implement.EscapeXmlReference");
    }


    public static Template velocityTemplate(String template){
        VelocityEngine ve = velocityEngine();
        ve.init();
        //取得velocity的模版
        return ve.getTemplate(template);
    }

    public static VelocityEngine velocityEngine(){
        VelocityEngine ve = new VelocityEngine(p);
        ve.init();
        //取得velocity的模版
        return ve;
    }

    public static Reader parse(Map<String, Object> map, String template){
        VelocityContext context = new VelocityContext(map);
        Template t = velocityTemplate(template);
        StringWriter sw = new StringWriter();
        t.merge(context, sw);
        return new StringReader(sw.toString());
    }
}
