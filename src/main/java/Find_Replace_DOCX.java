import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

//import org.apache.poi.xwpf.usermodel.XWPFTable;
//import org.apache.poi.xwpf.usermodel.XWPFTableCell;
//import org.apache.poi.xwpf.usermodel.XWPFTableRow;
/**
 * @author kishan.c.s 09-Feb-2016
 * Code to Find and replace content in docx
 */

public class Find_Replace_DOCX {

    public static void main(String[] args) throws Exception {
        new Find_Replace_DOCX().find_replace_in_DOCX();
    }

    public void find_replace_in_DOCX() throws Exception {
        try {

/**
 * if uploaded doc then use HWPF else if uploaded Docx file use
 * XWPFDocument
 */
            XWPFDocument doc = new XWPFDocument(
                    OPCPackage.open("VvT_Verklaring_van_Toepasselijkheid.docx"));
            Set<XWPFRun> occurences = new HashSet<>();
           // XWPFChart chart = generateDashboard(doc);
            XWPFParagraph p = findDashboardParagraph(doc);

            if (p != null) {
                XWPFChart chart = generateDashboard(doc);

                // CODE HIERONDER GEEFT DIE WORD ERRORS
//                XWPFRun run = p.createRun();

//                XmlCursor cursor = p.getCTP().newCursor();
//                XWPFParagraph newPara = doc.insertNewParagraph(cursor);
//                newPara.setAlignment (ParagraphAlignment.CENTER); // center
//                XWPFRun newParaRun = newPara.createRun();
//  /*              newParaRun.addPicture(new FileInputStream("./doc/bus.png"),XWPFDocument.PICTURE_TYPE_PNG,"bus.png,",Units.toEMU(200), Units.toEMU(200));
//                doc.removeBodyElement(doc.getPosOfParagraph(p));*/
//                newParaRun.addChart(doc.getRelationId(chart));
//                doc.removeBodyElement(doc.getPosOfParagraph(p));

               // run.addChart(doc.getRelationId(chart));
                // attach the chart here
             /*   java.lang.reflect.Method attach = XWPFChart.class.getDeclaredMethod("attach", String.class, XWPFRun.class);
                attach.setAccessible(true);
                attach.invoke(chart, doc.getRelationId(chart), run);
                chart.setChartBoundingBox(7*Units.EMU_PER_CENTIMETER, 7*Units.EMU_PER_CENTIMETER);*/
               // run.setText("tested");
                //run.addChart(doc.getRelationId(chart));
                // attach the chart here
          /*      java.lang.reflect.Method attach = XWPFChart.class.getDeclaredMethod("attach", String.class, XWPFRun.class);
                attach.setAccessible(true);
                //attach.invoke(chart, doc.getRelationId(chart), run);
                List<POIXMLDocumentPart> charts = doc.getRelations().stream().filter(poixmlDocumentPart -> poixmlDocumentPart instanceof XWPFChart).collect(Collectors.toList());
                attach.invoke(charts.get(0), doc.getRelationId(charts.get(0)), run);*/

/*

               XmlCursor cursor = p.getCTP().newCursor();

                XWPFParagraph newP = doc.createParagraph();
                newP.getCTP().setPPr(p.getCTP().getPPr());
                XWPFRun newR = newP.createRun();
                newR.getCTR().setRPr(p.getRuns().get(0).getCTR().getRPr());
                newR.setText("");
                newR.addChart(doc.getRelationId(chart));
                XmlCursor c2 = newP.getCTP().newCursor();
                c2.moveXml(cursor);
                c2.dispose();
                cursor.removeXml(); // Removes replacement text paragraph
                cursor.dispose();
*/
            }

/*            for (XWPFRun run : occurences) {
                run.addChart(doc.getRelationId(chart));
            }*/


            /*
             * for (XWPFTable tbl : doc.getTables()) { for (XWPFTableRow row :
             * tbl.getRows()) { for (XWPFTableCell cell : row.getTableCells()) {
             * for (XWPFParagraph p : cell.getParagraphs()) { for (XWPFRun r :
             * p.getRuns()) { String text = r.getText(0); if
             * (text.contains("needle")) { text = text.replace("key",
             * "kkk"); r.setText(text); } } } } } }
             */

           // XWPFChart chart = generateDashboard(doc);

            doc.write(new FileOutputStream("output.docx"));
        } finally {

        }

    }


    private XWPFParagraph findDashboardParagraph(XWPFDocument doc) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("#VVT.DASHBOARD#")) {
                        text = text
                                .replace("#VVT.DASHBOARD#", "");
                        r.setText(text, 0);

                        return p;
                        //occurences.add(r);
                    }
                }
            }
        }
        return null;
    }


    private XWPFChart generateDashboard(XWPFDocument doc) throws Exception {

        // try (XWPFDocument document = new XWPFDocument()) {


        double inch = 1_440;
        String[] headers = new String[]{"Norm", "Normnr", "Omschrijving", "Status"};


        // create the data
        String[] categories = new String[]{"[TEST] Nog in te vullen", "[TEST] Niet geïmplementeerd", "[TEST] Deels geïmplementeerd", "[TEST] Geïmplementeerd", "[TEST] Geaccepteerd risico"};
        XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories, "controls");
        Integer[] valuesA = new Integer[]{30, 20, 5, 70, 2};
        XDDFNumericalDataSource<Integer> val = XDDFDataSourcesFactory.fromArray(valuesA, "controls", 0);

        // Replace the existing chart instead of creating new chart
        // create the chart
        // XWPFChart chart = doc.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

        // Get the existing Pie3DChart
        XWPFChart chart=null;
        for (POIXMLDocumentPart part : doc.getRelations()) {
            if (part instanceof XWPFChart) {
                chart = (XWPFChart) part;
                break;
            }
        }
        chart.setTitleText("Controls 3");

        List<XDDFChartData> chartSeries = chart.getChartSeries();
        XDDFPie3DChartData data = (XDDFPie3DChartData) chartSeries.get(0);
        XDDFChartData.Series series = data.getSeries(0);

        // replace the existing data
        series.replaceData(cat, val);

        // create chart data
//        XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);

        // create series
        // if only one series do not vary colors for each bar
        //((XDDFBarChartData) data).setVaryColors(false);
//        XDDFChartData.Series series = data.addSeries(cat, val);
        // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
        // Excel cannot opening the workbook after creatingg that incomplete Table.
        // So updating the chart data in Word is not possible.
        //series.setTitle("a", chart.setSheetTitle("a", 1));
//        series.setTitle("Controls", setTitleInDataSheet(chart, "Controls", 1));
//        series.setShowLeaderLines(true);
			/*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
			*/

        // plot chart data
        chart.plot(data);

        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);
        legend.setOverlay(false);


        // attach the chart here
/*        java.lang.reflect.Method attach = XWPFChart.class.getDeclaredMethod("attach", String.class, XWPFRun.class);
        attach.setAccessible(true);
        attach.invoke(chart, document.getRelationId(chart), run);*/


        // Write the output to a file
/*        try (FileOutputStream fileOut = new FileOutputStream("VVTPIEDOC-v02.docx")) {
            doc.write(fileOut);
        }*/
        return chart;
    }

    // Methode to set title in the data sheet without creating a Table but using the sheet data only.
    // Creating a Table is not really necessary.
    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

}