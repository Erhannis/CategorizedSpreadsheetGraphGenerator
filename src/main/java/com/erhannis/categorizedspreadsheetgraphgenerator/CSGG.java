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
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFNoFillProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STCompoundLine;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineCap;
import org.openxmlformats.schemas.drawingml.x2006.main.STPenAlignment;

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
                    double sx = Double.parseDouble(line[xcol]);
                    double sy = Double.parseDouble(line[ycol]);

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
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 20, 30); ////

            XSSFChart chart = drawing.createChart(anchor);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
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

                series.setMarkerStyle(MarkerStyle.CIRCLE);
                series.setMarkerSize((short) 2);

                XDDFLineProperties lineProperties = new XDDFLineProperties();
                int hc = ("blahblahblah"+name+"blahblah").hashCode(); // Deriving color from name.  Plain hash code was too small
                double f = 0.8; // Scaling, so they don't conflict with the light background
                byte[] rgb = new byte[] {(byte)(((hc & 0x00FF0000) >>> 16)*f), (byte)(((hc & 0x0000FF00) >>> 8)*f), (byte)((hc & 0x000000FF)*f)};
                lineProperties.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(rgb)));
                lineProperties.setWidth(0.1);
                XDDFShapeProperties shapeProperties = series.getShapeProperties();
                if (shapeProperties == null) {
                    shapeProperties = new XDDFShapeProperties();
                }
                shapeProperties.setLineProperties(lineProperties);
                series.setShapeProperties(shapeProperties);
            }

            chart.plot(data);

            // --> https://stackoverflow.com/a/51541623
            // do not auto delete the title
            if (chart.getCTChart().getAutoTitleDeleted() == null) {
                chart.getCTChart().addNewAutoTitleDeleted();
            }
            chart.getCTChart().getAutoTitleDeleted().setVal(false);

            // plot area background and border line
            if (chart.getCTChartSpace().getSpPr() == null) {
                chart.getCTChartSpace().addNewSpPr();
            }
            if (chart.getCTChartSpace().getSpPr().getSolidFill() == null) {
                chart.getCTChartSpace().getSpPr().addNewSolidFill();
            }
            if (chart.getCTChartSpace().getSpPr().getSolidFill().getSrgbClr() == null) {
                chart.getCTChartSpace().getSpPr().getSolidFill().addNewSrgbClr();
            }
            chart.getCTChartSpace().getSpPr().getSolidFill().getSrgbClr().setVal(new byte[]{(byte) 255, (byte) 255, (byte) 255});
            if (chart.getCTChartSpace().getSpPr().getLn() == null) {
                chart.getCTChartSpace().getSpPr().addNewLn();
            }
            chart.getCTChartSpace().getSpPr().getLn().setW(Units.pixelToEMU(1));
            if (chart.getCTChartSpace().getSpPr().getLn().getSolidFill() == null) {
                chart.getCTChartSpace().getSpPr().getLn().addNewSolidFill();
            }
            if (chart.getCTChartSpace().getSpPr().getLn().getSolidFill().getSrgbClr() == null) {
                chart.getCTChartSpace().getSpPr().getLn().getSolidFill().addNewSrgbClr();
            }
            chart.getCTChartSpace().getSpPr().getLn().getSolidFill().getSrgbClr().setVal(new byte[]{(byte) 0, (byte) 0, (byte) 0});

            // line style of cat axis
            if (chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr() == null) {
                chart.getCTChart().getPlotArea().getCatAxArray(0).addNewSpPr();
            }
            if (chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn() == null) {
                chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().addNewLn();
            }
            chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().setW(Units.pixelToEMU(1));
            if (chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().getSolidFill() == null) {
                chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().addNewSolidFill();
            }
            if (chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().getSolidFill().getSrgbClr() == null) {
                chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().getSolidFill().addNewSrgbClr();
            }
            chart.getCTChart().getPlotArea().getCatAxArray(0).getSpPr().getLn().getSolidFill().getSrgbClr()
                    .setVal(new byte[]{(byte) 0, (byte) 0, (byte) 0});

            //line style of val axis
            if (chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr() == null) {
                chart.getCTChart().getPlotArea().getValAxArray(0).addNewSpPr();
            }
            if (chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn() == null) {
                chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().addNewLn();
            }
            chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().setW(Units.pixelToEMU(1));
            if (chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().getSolidFill() == null) {
                chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().addNewSolidFill();
            }
            if (chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().getSolidFill().getSrgbClr() == null) {
                chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().getSolidFill().addNewSrgbClr();
            }
            chart.getCTChart().getPlotArea().getValAxArray(0).getSpPr().getLn().getSolidFill().getSrgbClr()
                    .setVal(new byte[]{(byte) 0, (byte) 0, (byte) 0});
            // <-- https://stackoverflow.com/a/51541623

            //TODO I don't know how to get gridlines to work, and apparently nobody else does, either
            
//            CTShapeProperties ctShapeProperties = chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines().addNewSpPr();
//            CTLineProperties ctLineProperties1 = ctShapeProperties.addNewLn();
//            ctLineProperties1.setW(9525);
//            ctLineProperties1.setCap(STLineCap.FLAT);
//            ctLineProperties1.setCmpd(STCompoundLine.SNG);
//            ctLineProperties1.setAlgn(STPenAlignment.CTR);
            
//            chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines().addNewSpPr().addNewSolidFill();
//            chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines().addNewSpPr().addNewSolidFill();

//            chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines();
//            chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines();

//            CTPlotArea plotArea = chart.getCTChart().getPlotArea();
//            plotArea.getCatAxArray()[0].addNewMajorGridlines();
//            plotArea.getValAxArray()[0].addNewMajorGridlines();

            return wb;
        } finally {
            reset();
        }
    }
}
