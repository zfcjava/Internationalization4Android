package com.canjun.onenine;

import org.jdom2.Attribute;
import org.jdom2.Comment;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;


import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class AndroidString {

    /**
     * 生成Android中中字符串文件
     * @param map 数据源
     * @param destFile 写入的目标文件地址
     */
    public static void writeXml(Map<String, String> map, String destFile) {
        //创建文档
        Document document = new Document();
        //创建根元素
        Element people = new Element("resources");
        //把根元素加入到document中
        document.addContent(people);

        //创建注释
        Comment rootComment = new Comment("将数据从程序输出到XML中！");
        people.addContent(rootComment);

        for (Map.Entry<String,String> e: map.entrySet()) {
            //创建父元素
            Element element = new Element("string");
            //把元素加入到根元素中
            people.addContent(element);
            //设置person1元素属性
            element.setAttribute("name", e.getKey());
            element.addContent(e.getValue());
        }


        //设置xml输出格式
        Format format = Format.getPrettyFormat();
        format.setEncoding("utf-8");//设置编码
        format.setIndent("    ");//设置缩进


        //得到xml输出流
        XMLOutputter out = new XMLOutputter(format);
        //把数据输出到xml中
        try {
            out.output(document, new FileOutputStream(destFile));//或者FileWriter
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
