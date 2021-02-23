import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.XDDFFillProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;

public class PieChartWordXDDF {

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

    public static void main(String[] args) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {

            // create the data
            String[] categories = new String[] { "Nog in te vullen", "Niet geïmplementeerd", "Deels geïmplementeerd","Geïmplementeerd","Geaccepteerd risico"};
            XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories,"controls");
            Integer[] valuesA = new Integer[] { 30, 20, 5,70,2};
            XDDFNumericalDataSource<Integer> val = XDDFDataSourcesFactory.fromArray(valuesA,"controls",0);


            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
            chart.setTitleText("Controls");

            // create chart data
            XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);

            // create series
            // if only one series do not vary colors for each bar
            //((XDDFBarChartData) data).setVaryColors(false);
            XDDFChartData.Series series = data.addSeries(cat, val);
            // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
            // Excel cannot opening the workbook after creatingg that incomplete Table.
            // So updating the chart data in Word is not possible.
            //series.setTitle("a", chart.setSheetTitle("a", 1));
            series.setTitle("Controls", setTitleInDataSheet(chart, "Controls", 1));
            series.setShowLeaderLines(true);
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


            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("PieChartDOC.docx")) {
                document.write(fileOut);
            }
        }
    }
}