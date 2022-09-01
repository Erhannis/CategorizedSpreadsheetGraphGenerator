/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.erhannis.categorizedspreadsheetgraphgenerator;

import com.martiansoftware.jsap.FlaggedOption;
import com.martiansoftware.jsap.JSAP;
import com.martiansoftware.jsap.JSAPException;
import com.martiansoftware.jsap.JSAPResult;
import com.martiansoftware.jsap.Parameter;
import com.martiansoftware.jsap.SimpleJSAP;
import com.martiansoftware.jsap.Switch;
import com.martiansoftware.jsap.UnflaggedOption;
import com.opencsv.exceptions.CsvValidationException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFNoFillProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Erhannis
 */
public class Main {

    public static void main(String[] args) throws IOException, JSAPException, CsvValidationException {
        SimpleJSAP jsap = new SimpleJSAP(
                "CSGG",
                "Categorized Spreadsheet Graph Generator - pipe csv data to it, with cols for series name, x, and y, and it will generate a corresponding Excel scatter plot.",
                new Parameter[]{
                    new Switch("help2", 'h', null, "Print help."),
                    new Switch("dieonerror", 'q', "dieonerror", "Quit if problem occurs during parsing.  Default is to skip erroring rows."),
                    new FlaggedOption("skipn", JSAP.INTEGER_PARSER, "0", JSAP.NOT_REQUIRED, 's', "skipn", "Skip N rows (e.g. headers)"),
                    new FlaggedOption("categorycol", JSAP.INTEGER_PARSER, "0", JSAP.NOT_REQUIRED, 'c', "categorycol", "Column of category (0 is first column)"),
                    new FlaggedOption("xcol", JSAP.INTEGER_PARSER, "1", JSAP.NOT_REQUIRED, 'x', "xcol", "Column of x values (0 is first column)"),
                    new FlaggedOption("ycol", JSAP.INTEGER_PARSER, "2", JSAP.NOT_REQUIRED, 'y', "ycol", "Column of y values (0 is first column)")
                }
        );

        JSAPResult config = jsap.parse(args);
        if (jsap.messagePrinted()) {
            System.exit(1);
        }
        if (config.getBoolean("help2")) {
            System.out.println(jsap.getHelp());
            System.exit(0);
        }

        boolean dieonerror = config.getBoolean("dieonerror");
        int skipn = config.getInt("skipn");
        int categorycol = config.getInt("categorycol");
        int xcol = config.getInt("xcol");
        int ycol = config.getInt("ycol");


        XSSFWorkbook result = new CSGG(dieonerror, skipn, categorycol, xcol, ycol).parse(System.in);
        //XSSFWorkbook result = new CSGG(dieonerror, skipn, categorycol, xcol, ycol).parse(new FileInputStream("test.csv"));
        System.out.println("result: " + result);
        try ( FileOutputStream fileOut = new FileOutputStream("test_"+System.currentTimeMillis()+".xlsx")) {
            result.write(fileOut);
        }
        //TODO Do you need to close the workbook?
    }
}
