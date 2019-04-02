package com.canjun.onenine;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class Main {


    public static HashMap<String, String> map = new HashMap<String, String>();

    public static void main(String[] args) {
	// write your code here

        processFile("/Users/zhngfengcheng/trans/file.xlsx","/Users/zhngfengcheng/trans/strings.xml");
    }




    private static void processFile(String filePath,String destFile){

        File xlsxFile = new File(filePath);
        if (!xlsxFile.exists()) {
            System.err.println("Not found or not a file: " + xlsxFile.getPath());
            return;
        }

        //暂时显示第一列和第二列
        int[] columIndexs = {1,2};

        // The package open is instantaneous, as it should be.
        OPCPackage p = null;
        try {
            p = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ);
            XLSX2CSV xlsx2csv = new XLSX2CSV(p, System.out,map, 2, columIndexs);
            try {
                xlsx2csv.process();
                System.out.println(map);

                AndroidString.writeXml(map, destFile);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (OpenXML4JException e) {
                e.printStackTrace();
            } catch (ParserConfigurationException e) {
                e.printStackTrace();
            } catch (SAXException e) {
                e.printStackTrace();
            }
            try {
                p.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }


}
