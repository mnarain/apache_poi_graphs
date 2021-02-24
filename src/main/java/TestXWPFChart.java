import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class TestXWPFChart {

    public static void main(String[] args) throws Exception {
FileInputStream inpuFile=new FileInputStream("dashboard.docx");
        FileOutputStream outFile = new FileOutputStream("output.docx");
        @SuppressWarnings("resource")
        XWPFDocument document = new XWPFDocument(inpuFile);
        XWPFChart chart=null;
        for (POIXMLDocumentPart part : document.getRelations()) {
            if (part instanceof XWPFChart) {
                chart = (XWPFChart) part;
                break;
            }
        }
        //change chart title from "Chart Title" to XWPF CHART
        CTChart ctChart = chart.getCTChart();
        CTTitle title = ctChart.getTitle();
        CTTx tx = title.addNewTx();
        CTTextBody rich = tx.addNewRich();
        rich.addNewBodyPr();
        rich.addNewLstStyle();
        CTTextParagraph p = rich.addNewP();
        CTRegularTextRun r = p.addNewR();
        r.addNewRPr();
        r.setT("XWPF CHART");

        //write modified chart in output docx file
        document.write(outFile);
}
}