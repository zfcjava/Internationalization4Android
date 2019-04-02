package com.canjun.onenine;


import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
 
/* ====================================================================
        Licensed to the Apache Software Foundation (ASF) under one or more
        contributor license agreements.  See the NOTICE file distributed with
        this work for additional information regarding copyright ownership.
        The ASF licenses this file to You under the Apache License, Version 2.0
        (the "License"); you may not use this file except in compliance with
        the License.  You may obtain a copy of the License at
        http://www.apache.org/licenses/LICENSE-2.0
        Unless required by applicable law or agreed to in writing, software
        distributed under the License is distributed on an "AS IS" BASIS,
        WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
        See the License for the specific language governing permissions and
        limitations under the License.
        ==================================================================== */
 
 
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.HashMap;


import javax.xml.parsers.ParserConfigurationException;
 
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;



/**
 * A rudimentary XLSX -> CSV processor modeled on the
 * POI sample program XLS2CSVmra from the package
 * org.apache.poi.hssf.eventusermodel.examples.
 * As with the HSSF version, this tries to spot missing
 * rows and cells, and output empty entries for them.
 * <p/>
 * Data sheets are read using a SAX parser to keep the
 * memory footprint relatively small, so this should be
 * able to read enormous workbooks.  The styles table and
 * the shared-string table must be kept in memory.  The
 * standard POI styles table class is used, but a custom
 * (read-only) class is used for the shared string table
 * because the standard POI SharedStringsTable grows very
 * quickly with the number of unique strings.
 * <p/>
 * For a more advanced implementation of SAX event parsing
 * of XLSX files, see {@link XSSFEventBasedExcelExtractor}
 * and {@link XSSFSheetXMLHandler}. Note that for many cases,
 * it may be possible to simply use those with a custom
 * {@link SheetContentsHandler} and no SAX code needed of
 * your own!
 */
public class XLSX2CSV {


    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     * of parsing the Sheet XML, and outputs the contents
     * as a (basic) CSV.
     */
    private class SheetToCSV implements SheetContentsHandler {
        private boolean firstCellOfRow = false;
        private int currentRow = -1;
        private int currentCol = -1;
 
        private void outputMissingRows(int number) {
            for (int i = 0; i < number; i++) {
                for (int j = 0; j < minColumns; j++) {
                    output.print(',');
                }
                output.print('\n');
            }
        }

        public void startRow(int rowNum) {
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum - currentRow - 1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }
 

        public void endRow(int rowNum) {
            // Ensure the minimum number of columns
            //TODO 暂时不需要补全
//            for (int i = currentCol; i < minColumns; i++) {
//                output.print(',');
//            }
            output.print('\n');
        }
 

        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            }
//            else {
//                output.print(',');
//            }
 
            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }
 
            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            if (isTargetIndex(thisCol)) {
//                int missedCols = thisCol - currentCol - 1;
//                for (int i = 0; i < missedCols; i++) {
//                    output.print(',');
//                }
//                currentCol = thisCol;
                if (!isFirstIndex(thisCol)) {
                    output.print(',');

                    preValue = formattedValue;


                    map.put(preKey, preValue);
                    preValue = "";
                    preKey = "";
                } else {
                    preKey = formattedValue;
                }

                // Number or string?
                try {
                    Double.parseDouble(formattedValue);
                    output.print(formattedValue);
                } catch (NumberFormatException e) {
                    output.print('"');
                    output.print(formattedValue);
                    output.print('"');
                }
            }

        }
 
    
    }
 
 
    ///////////////////////////////////////
 
    private final OPCPackage xlsxPackage;
 
    /**
     * Number of columns to read starting with leftmost
     */
    private final int minColumns;

    /**
     * 指定输入的列索引
     */
    private final int[] columsIndexs;
 
    /**
     * Destination for data
     */
    private final PrintStream output;


    private String preKey = "";

    private String preValue = "";


    /**
     * 将数据以键值对的形式存储到内存中
     */
    private HashMap<String, String> map;
 
    /**
     * Creates a new XLSX -> CSV converter
     *
     * @param pkg         The XLSX package to process
     * @param output      The PrintStream to output the CSV to
     * @param minColumns  The minimum number of columns to output, or -1 for no minimum
     * @param columsIndexs 选择置顶要打印的行索引
     */
    public XLSX2CSV(OPCPackage pkg, PrintStream output, int minColumns,int[]columsIndexs) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
        this.columsIndexs = columsIndexs;
    }

    /**
     * Creates a new XLSX -> CSV converter
     *
     * @param pkg         The XLSX package to process
     * @param output      The PrintStream to output the CSV to
     * @param minColumns  The minimum number of columns to output, or -1 for no minimum
     * @param columsIndexs 选择置顶要打印的行索引
     */
    public XLSX2CSV(OPCPackage pkg, PrintStream output, HashMap<String,String> map,int minColumns, int[]columsIndexs) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
        this.columsIndexs = columsIndexs;
        this.map = map;
    }



    /**
     * 根据索引值判断当前列是否需要打印
     * @param index
     * @return
     */
    public boolean isTargetIndex(int index) {
        for (int i = 0; i < columsIndexs.length; i++) {
            if (index == columsIndexs[i]) {
                return true;
            }
        }
        return false;
    }


    /**
     * 第一个属性值前面不增加","
     * @param index
     * @return
     */
    public boolean isFirstIndex(int index) {
        return columsIndexs.length > 0 && columsIndexs[0] == index;
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles
     * @param strings
     * @param sheetInputStream
     */
    public void processSheet(
            StylesTable styles,
            ReadOnlySharedStringsTable strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser =org.apache.poi.ooxml.util.SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }
 
    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    public void process()
            throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        while (iter.hasNext()) {
            InputStream stream = iter.next();
            String sheetName = iter.getSheetName();
            this.output.println();
            this.output.println(sheetName + " [index=" + index + "]:");
            processSheet(styles, strings, new SheetToCSV(), stream);
            stream.close();
            ++index;
        }
    }

}
