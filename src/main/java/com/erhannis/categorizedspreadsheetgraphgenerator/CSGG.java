/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.erhannis.categorizedspreadsheetgraphgenerator;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.HashMap;
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
public class CSGG {
    private static class SheetInfo {
        public XSSFSheet sheet;
        public int rows = 0;
    }
    
    public final boolean dieonerror;
    public final int skipn;
    public final int categorycol;
    public final int xcol;
    public final int ycol;
    
    public CSGG(boolean dieonerror, int skipn, int categorycol, int xcol, int ycol) {
        this.dieonerror = dieonerror;
        this.skipn = skipn;
        this.categorycol = categorycol;
        this.xcol = xcol;
        this.ycol = ycol;
    }

    private XSSFWorkbook wb;
    private HashMap<String, SheetInfo> sheets;
    
    private SheetInfo getSheet(String name) {
        if (sheets.containsKey(name)) {
            return sheets.get(name);
        } else {
            XSSFSheet sheet = wb.createSheet(name);
            SheetInfo si = new SheetInfo();
            si.sheet = sheet;
            sheets.put(name, si);
            return si;
        }
    }
    
    private void reset() {
        System.out.println("resetting...");
        wb = null;
        sheets = null;
    }
    
    // A lot of this is derived from https://stackoverflow.com/a/59067102
    public XSSFWorkbook parse(InputStream in) throws IOException, CsvValidationException {
        reset();
        
        CSVReader csv = new CSVReader(new InputStreamReader(in));
        for (int i = 0; i < skipn; i++) {
            csv.readNextSilently();
        }
        
        try {
            this.wb = new XSSFWorkbook();
            this.sheets = new HashMap<String, SheetInfo>();
            
            XSSFSheet chartSheet = wb.createSheet("Chart");
            
            String[] line = null;
            while ((line = csv.readNext()) != null) {
                try {
                    String cat = line[categorycol];
                    String sx = line[xcol];
                    String sy = line[ycol];

                    SheetInfo si = getSheet(cat);
                    Row row = si.sheet.createRow((short) si.rows);
                    si.rows++;
                    Cell x = row.createCell(0);
                    x.setCellValue(sx);
                    Cell y = row.createCell(1);
                    y.setCellValue(sy);
                } catch (Throwable t) {
                    if (dieonerror) {
                        throw t;
                    } else {
                        t.printStackTrace();
                    }
                }
            }
            System.out.println("Input has ended");
            
            XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15); ////

            XSSFChart chart = drawing.createChart(anchor);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            XDDFValueAxis bottomAxis = chart.createValueAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("x"); // https://stackoverflow.com/questions/32010765
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("y");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
            
            for (String name : sheets.keySet()) {
                SheetInfo si = sheets.get(name);
                
                XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(si.sheet, new CellRangeAddress(0, si.rows-1, 0, 0));
                XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(si.sheet, new CellRangeAddress(0, si.rows-1, 1, 1));
                
                XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(xs, ys);
                series.setTitle(name, null); // https://stackoverflow.com/questions/21855842
                series.setSmooth(false); // https://stackoverflow.com/questions/39636138

                series.setMarkerStyle(MarkerStyle.DOT);
                series.setMarkerSize((short) 2);

                XDDFLineProperties lineProperties = new XDDFLineProperties();
                lineProperties.setWidth(0.0);
                XDDFShapeProperties shapeProperties = series.getShapeProperties();
                if (shapeProperties == null) {
                    shapeProperties = new XDDFShapeProperties();
                }
                shapeProperties.setLineProperties(lineProperties);
                series.setShapeProperties(shapeProperties);
            }

            chart.plot(data);

            return wb;
        } finally {
            reset();
        }
    }
}
