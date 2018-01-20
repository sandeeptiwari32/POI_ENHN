import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartLines;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTExtension;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumFmt;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarGrouping;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrossBetween;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLblAlgn;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickMark;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBodyProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.main.STCompoundLine;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineCap;
import org.openxmlformats.schemas.drawingml.x2006.main.STPenAlignment;
import org.openxmlformats.schemas.drawingml.x2006.main.STSchemeColorVal;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAnchoringType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextStrikeType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextUnderlineType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextVertOverflowType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextVerticalType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextWrappingType;

public class TestChart {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		@SuppressWarnings("resource")
		XWPFDocument document=new XWPFDocument();
        System.out.println(XWPFDocument.class.getResource("XWPFDocument.class"));
		FileOutputStream outStream = new FileOutputStream("dynamicChart.docx");
		XWPFChart chart =document.createChart(500000, 500000);
		XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        Row row;
        Cell cell;
        row = sheet.createRow(0);
        row.createCell(0);
        row.createCell(1).setCellValue("HEADER 1");

        for (int r = 1; r < 5; r++) {
            row = sheet.createRow(r);
            cell = row.createCell(0);
            cell.setCellValue("Series " + r);
            cell = row.createCell(1);
            cell.setCellValue(r+1);
        }
        addChart(chart,wb,true);
		XWPFChart chart1 =document.createChart();
		addChart(chart1, wb,false);
        document.write(outStream);
        System.out.println("done");
	}
	
	public static void addChart(XWPFChart chart,XSSFWorkbook wb,boolean isBar) throws IOException, InvalidFormatException
	{
		chart.setChartBoundingBox(5000000, 5000000);
		chart.setChartMargin(100, 100, 100, 100);
		chart.saveWorkbook(wb);
		CTChartSpace ctChartSpace = chart.getCTChartSpace();
		ctChartSpace.addNewDate1904().setVal(false);
		ctChartSpace.addNewLang().setVal("en-US");
		ctChartSpace.addNewRoundedCorners().setVal(false);
		
        CTChart ctChart = chart.getCTChart();
        
        CTTitle ctTitle = ctChart.addNewTitle();
        ctTitle.addNewOverlay().setVal(false);
        
        CTShapeProperties ctSpPr = ctTitle.addNewSpPr();
        ctSpPr.addNewNoFill();
        ctSpPr.addNewLn().addNewNoFill();
        ctSpPr.addNewEffectLst();
        
        CTTextBody ctTxPr = ctTitle.addNewTxPr();
        
        CTTextBodyProperties ctBodyPr = ctTxPr.addNewBodyPr();
        ctBodyPr.setRot(0);
        ctBodyPr.setSpcFirstLastPara(true);
        ctBodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
        ctBodyPr.setVert(STTextVerticalType.HORZ);
        ctBodyPr.setWrap(STTextWrappingType.SQUARE);
        ctBodyPr.setAnchor(STTextAnchoringType.CTR);
        ctBodyPr.setAnchorCtr(true);
        ctTxPr.addNewLstStyle();
        CTTextParagraph ctP = ctTxPr.addNewP();
        
        CTTextCharacterProperties ctRPr = ctP.addNewPPr().addNewDefRPr();
        ctRPr.setSz(1400);
        ctRPr.setKern(1200);
        ctRPr.setI(false);
        ctRPr.setB(false);
        ctRPr.setU(STTextUnderlineType.NONE);
        ctRPr.setStrike(STTextStrikeType.NO_STRIKE);
        ctRPr.setBaseline(0);
        ctRPr.setSpc(0);
        
        CTSchemeColor ctSchemeClr = ctRPr.addNewSolidFill().addNewSchemeClr();
        ctSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctSchemeClr.addNewLumMod().setVal(65000);
        ctSchemeClr.addNewLumOff().setVal(35000);
        
        ctRPr.addNewLatin().setTypeface("+mn-lt");
        ctRPr.addNewEa().setTypeface("+mn-ea");
        ctRPr.addNewCs().setTypeface("+mn-cs");
        
        ctP.addNewEndParaRPr().setLang("en-US");
        
        ctChart.addNewAutoTitleDeleted().setVal(false);
        
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        ctPlotArea.addNewLayout();
        
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        if(isBar)
        	ctBarChart.addNewBarDir().setVal(STBarDir.BAR);
        else
        	ctBarChart.addNewBarDir().setVal(STBarDir.COL);
        ctBarChart.addNewGrouping().setVal(STBarGrouping.CLUSTERED);
        ctBarChart.addNewVaryColors().setVal(true);
        
        for (int r = 2; r < 3; r++) {
           CTBarSer ctBarSer = ctBarChart.addNewSer();
           ctBarSer.addNewIdx().setVal(r-2);
           ctBarSer.addNewOrder().setVal(r-2);
           
           CTSerTx ctSerTx = ctBarSer.addNewTx();
           
           CTStrRef ctStrRef = ctSerTx.addNewStrRef();
           ctStrRef.setF("Sheet1!$B$" + 1);  
           
           CTStrData ctStrCache = ctStrRef.addNewStrCache();
           ctStrCache.addNewPtCount().setVal(r-1);;
           CTStrVal ctStrCacheP = ctStrCache.addNewPt();
           ctStrCacheP.setIdx(r-2);
           ctStrCacheP.setV("HEADER 1");
           
           CTShapeProperties ctBarSpPr = ctBarSer.addNewSpPr();
           
           CTSchemeColor ctBarSchemeClr = ctBarSpPr.addNewSolidFill().addNewSchemeClr();
           ctBarSchemeClr.setVal(STSchemeColorVal.ACCENT_1);
           
           ctBarSpPr.addNewLn().addNewNoFill();
           ctBarSpPr.addNewEffectLst();
           
           ctBarSer.addNewInvertIfNegative().setVal(false);
           
           CTDLbls ctDlbls = ctBarSer.addNewDLbls();
           
           CTShapeProperties ctDlblsSpPr = ctDlbls.addNewSpPr() ;
           ctDlblsSpPr.addNewNoFill();
           ctDlblsSpPr.addNewLn().addNewNoFill();
           ctDlblsSpPr.addNewEffectLst();
           
           CTTextBody ctDlblsTxPr = ctDlbls.addNewTxPr();
           
           CTTextBodyProperties ctDlblsBodyPr = ctDlblsTxPr.addNewBodyPr();
           ctDlblsBodyPr.setRot(0);
           ctDlblsBodyPr.setSpcFirstLastPara(true);
           ctDlblsBodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
           ctDlblsBodyPr.setVert(STTextVerticalType.HORZ);
           ctDlblsBodyPr.setWrap(STTextWrappingType.SQUARE);
           ctDlblsBodyPr.setLIns(38100);
           ctDlblsBodyPr.setTIns(19050);
           ctDlblsBodyPr.setRIns(38100);
           ctDlblsBodyPr.setBIns(19050);
           ctDlblsBodyPr.setAnchor(STTextAnchoringType.CTR);
           ctDlblsBodyPr.setAnchorCtr(true);
           ctDlblsBodyPr.addNewSpAutoFit();
           
           ctDlblsTxPr.addNewLstStyle();
           
           CTTextParagraph ctDlblsP = ctDlblsTxPr.addNewP();
           
           CTTextCharacterProperties ctDlblsRPr = ctDlblsP.addNewPPr().addNewDefRPr();
           ctDlblsRPr.setSz(900);
           ctDlblsRPr.setKern(1200);
           ctDlblsRPr.setI(true);
           ctDlblsRPr.setB(true);
           ctDlblsRPr.setU(STTextUnderlineType.NONE);
           ctDlblsRPr.setStrike(STTextStrikeType.NO_STRIKE);
           ctDlblsRPr.setBaseline(0);
           
           CTSchemeColor ctDlblsSchemeClr = ctDlblsRPr.addNewSolidFill().addNewSchemeClr();
           ctDlblsSchemeClr.setVal(STSchemeColorVal.TX_1);
           ctDlblsSchemeClr.addNewLumMod().setVal(75000);
           ctDlblsSchemeClr.addNewLumOff().setVal(25000);
           
           ctDlblsRPr.addNewLatin().setTypeface("+mn-lt");
           ctDlblsRPr.addNewEa().setTypeface("+mn-ea");
           ctDlblsRPr.addNewCs().setTypeface("+mn-cs");
           
           ctDlblsP.addNewEndParaRPr().setLang("en-US");
           
           ctDlbls.addNewDLblPos().setVal(STDLblPos.OUT_END);
           
           ctDlbls.addNewShowLegendKey().setVal(true);
           
           ctDlbls.addNewShowVal().setVal(true);
           
           ctDlbls.addNewShowCatName().setVal(true);
           
           ctDlbls.addNewShowSerName().setVal(true);
           
           ctDlbls.addNewShowPercent().setVal(true);
           
           ctDlbls.addNewShowBubbleSize().setVal(false);
           
           ctDlbls.addNewShowLeaderLines().setVal(false);;
           
           CTExtension ctDlblsExt = ctDlbls.addNewExtLst().addNewExt();
           ctDlblsExt.setUri("{CE6537A1-D6FC-4f65-9D91-7224C49458BB}");
           ctDlblsExt.selectAttribute("xmlns:c15","http://schemas.microsoft.com/office/drawing/2012/chart");
           
           CTStrRef ctCatStrRef = ctBarSer.addNewCat().addNewStrRef();
           ctCatStrRef.setF("Sheet1!$A$2:$A$5");
           
           CTStrData ctCatStrCache = ctCatStrRef.addNewStrCache();
           ctCatStrCache.addNewPtCount().setVal(4);
           
           CTStrVal ctCachePt = ctCatStrCache.addNewPt();
           ctCachePt.setIdx(0);
           ctCachePt.setV("series 1");
           
           CTStrVal ctCachePt2 = ctCatStrCache.addNewPt();
           ctCachePt2.setIdx(1);
           ctCachePt2.setV("series 2");
           
           CTStrVal ctCachePt3 = ctCatStrCache.addNewPt();
           ctCachePt3.setIdx(2);
           ctCachePt3.setV("series 3");
           
           CTStrVal ctCachePt4 = ctCatStrCache.addNewPt();
           ctCachePt4.setIdx(3);
           ctCachePt4.setV("series 4");
           
           CTNumRef ctNumRef = ctBarSer.addNewVal().addNewNumRef();
           ctNumRef.setF("Sheet1!$B$2:$B$5");
           
           CTNumData ctNumCache = ctNumRef.addNewNumCache();
           ctNumCache.setFormatCode("General");
           
           ctNumCache.addNewPtCount().setVal(4);
           
           CTNumVal ctCacheNumPt = ctNumCache.addNewPt();
           ctCacheNumPt.setIdx(0);
           ctCacheNumPt.setV("2");
           
           CTNumVal ctCacheNumPt2 = ctNumCache.addNewPt();
           ctCacheNumPt2.setIdx(1);
           ctCacheNumPt2.setV("3");
           
           CTNumVal ctCacheNumPt3 = ctNumCache.addNewPt();
           ctCacheNumPt3.setIdx(2);
           ctCacheNumPt3.setV("4");
           
           CTNumVal ctCacheNumPt4 = ctNumCache.addNewPt();
           ctCacheNumPt4.setIdx(3);
           ctCacheNumPt4.setV("5");
           
           CTExtension valExt = ctBarSer.addNewExtLst().addNewExt();
           valExt.setUri("{C3380CC4-5D6E-409C-BE32-E72D297353CC}");
           valExt.selectAttribute("xmlns:c16","http://schemas.microsoft.com/office/drawing/2014/chart");
           
           CTDLbls CtBarDLbls = ctBarChart.addNewDLbls();
           
           CtBarDLbls.addNewDLblPos().setVal(STDLblPos.OUT_END);
           
           CtBarDLbls.addNewShowLegendKey().setVal(true);
           
           CtBarDLbls.addNewShowVal().setVal(true);
           
           CtBarDLbls.addNewShowCatName().setVal(true);
           
           CtBarDLbls.addNewShowSerName().setVal(true);
           
           CtBarDLbls.addNewShowPercent().setVal(true);
           
           CtBarDLbls.addNewShowBubbleSize().setVal(false);
           
           ctBarChart.addNewGapWidth().setVal(200);

        } 

        //telling the BarChart that it has axes and giving them Ids
        ctBarChart.addNewAxId().setVal(123456);
        ctBarChart.addNewAxId().setVal(123457);

        //cat axis
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx(); 
        ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        
        CTNumFmt ctNumFmt = ctCatAx.addNewNumFmt();
        ctNumFmt.setFormatCode("General");
        ctNumFmt.setSourceLinked(true);
        
        ctCatAx.addNewMajorTickMark().setVal(STTickMark.NONE);
        ctCatAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        
        CTShapeProperties ctCatSpPr = ctCatAx.addNewSpPr();
        ctCatSpPr.addNewNoFill();
        CTLineProperties catLn = ctCatSpPr.addNewLn();
        catLn.setW(9525);
        catLn.setCap(STLineCap.FLAT);
        catLn.setCmpd(STCompoundLine.SNG);
        catLn.setAlgn(STPenAlignment.CTR);
        
        CTSchemeColor ctCatSchemeClr = catLn.addNewSolidFill().addNewSchemeClr();
        ctCatSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctCatSchemeClr.addNewLumMod().setVal(15000);
        ctCatSchemeClr.addNewLumOff().setVal(85000);
        
        catLn.addNewRound();
        
        ctCatSpPr.addNewEffectLst();
        
        CTTextBody ctCatTxPr = ctCatAx.addNewTxPr();
        CTTextBodyProperties ctCatBodyPr = ctCatTxPr.addNewBodyPr();
        ctCatBodyPr.setRot(-60000000);
        ctCatBodyPr.setSpcFirstLastPara(true);
        ctCatBodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
        ctCatBodyPr.setVert(STTextVerticalType.HORZ);
        ctCatBodyPr.setWrap(STTextWrappingType.SQUARE);
        ctCatBodyPr.setAnchor(STTextAnchoringType.CTR);
        ctCatBodyPr.setAnchorCtr(true);
        ctCatTxPr.addNewLstStyle();
        CTTextParagraph ctCatP = ctCatTxPr.addNewP();
        
        CTTextCharacterProperties ctCatRPr = ctCatP.addNewPPr().addNewDefRPr();
        ctCatRPr.setSz(900);
        ctCatRPr.setKern(1200);
        ctCatRPr.setI(false);
        ctCatRPr.setB(false);
        ctCatRPr.setU(STTextUnderlineType.NONE);
        ctCatRPr.setStrike(STTextStrikeType.NO_STRIKE);
        ctCatRPr.setBaseline(0);
        
        CTSchemeColor ctCatPSchemeClr = ctCatRPr.addNewSolidFill().addNewSchemeClr();
        ctCatPSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctCatPSchemeClr.addNewAlphaMod().setVal(65000);
        ctCatPSchemeClr.addNewAlphaOff().setVal(35000);
        
        ctCatRPr.addNewLatin().setTypeface("+mn-lt");
        ctCatRPr.addNewEa().setTypeface("+mn-ea");
        ctCatRPr.addNewCs().setTypeface("+mn-cs");
        
        ctCatP.addNewEndParaRPr().setLang("en-US");
        
        ctCatAx.addNewCrossAx().setVal(123457);
        
        ctCatAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO);
        
        ctCatAx.addNewAuto().setVal(true);
        
        ctCatAx.addNewLblAlgn().setVal(STLblAlgn.CTR);
        
        ctCatAx.addNewLblOffset().setVal(100);
        
        ctCatAx.addNewNoMultiLvlLbl().setVal(false);
        
        //val axis
        CTValAx ctValAx = ctPlotArea.addNewValAx(); 
        ctValAx.addNewAxId().setVal(123457); //id of the val axis
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        
        CTChartLines ctValMajorGridLine = ctValAx.addNewMajorGridlines();
        
        CTShapeProperties ctValMajorSpPr = ctValMajorGridLine.addNewSpPr();
        
        CTLineProperties catValLn = ctValMajorSpPr.addNewLn();
        catValLn.setW(9525);
        catValLn.setCap(STLineCap.FLAT);
        catValLn.setCmpd(STCompoundLine.SNG);
        catValLn.setAlgn(STPenAlignment.CTR);
        
        CTSchemeColor ctValSchemeClr = catValLn.addNewSolidFill().addNewSchemeClr();
        ctValSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctValSchemeClr.addNewLumMod().setVal(15000);
        ctValSchemeClr.addNewLumOff().setVal(85000);
        
        catValLn.addNewRound();
        
        ctValMajorSpPr.addNewEffectLst();
        
        CTNumFmt ctValNumFmt = ctValAx.addNewNumFmt();
        ctValNumFmt.setFormatCode("General");
        ctValNumFmt.setSourceLinked(true);
        
        ctValAx.addNewMajorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        
        CTShapeProperties ctValSpPr = ctValAx.addNewSpPr();
        ctValSpPr.addNewNoFill();
        ctValSpPr.addNewLn().addNewNoFill();
        ctValSpPr.addNewEffectLst();
        
        CTTextBody ctValTxPr = ctValAx.addNewTxPr();
        CTTextBodyProperties ctValBodyPr = ctValTxPr.addNewBodyPr();
        ctValBodyPr.setRot(-60000000);
        ctValBodyPr.setSpcFirstLastPara(true);
        ctValBodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
        ctValBodyPr.setVert(STTextVerticalType.HORZ);
        ctValBodyPr.setWrap(STTextWrappingType.SQUARE);
        ctValBodyPr.setAnchor(STTextAnchoringType.CTR);
        ctValBodyPr.setAnchorCtr(true);
        ctValTxPr.addNewLstStyle();
        
        CTTextParagraph ctValP = ctValTxPr.addNewP();
        
        CTTextCharacterProperties ctValDefRPr = ctValP.addNewPPr().addNewDefRPr();
        ctValDefRPr.setSz(900);
        ctValDefRPr.setKern(1200);
        ctValDefRPr.setI(false);
        ctValDefRPr.setB(false);
        ctValDefRPr.setU(STTextUnderlineType.NONE);
        ctValDefRPr.setStrike(STTextStrikeType.NO_STRIKE);
        ctValDefRPr.setBaseline(0);
        
        CTSchemeColor ctValPSchemeClr = ctValDefRPr.addNewSolidFill().addNewSchemeClr();
        ctValPSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctValPSchemeClr.addNewLumMod().setVal(65000);
        ctValPSchemeClr.addNewLumOff().setVal(35000);
        
        ctValDefRPr.addNewLatin().setTypeface("+mn-lt");
        ctValDefRPr.addNewEa().setTypeface("+mn-ea");
        ctValDefRPr.addNewCs().setTypeface("+mn-cs");
        
        ctValP.addNewEndParaRPr().setLang("en-US");
        
       
        ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
        ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO);
        ctValAx.addNewCrossBetween().setVal(STCrossBetween.BETWEEN);
        
        
        CTShapeProperties ctPlotSpPr = ctPlotArea.addNewSpPr();
        ctPlotSpPr.addNewNoFill();
        ctPlotSpPr.addNewLn().addNewNoFill();
        ctPlotSpPr.addNewEffectLst();
        
        //legend
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);
        
        CTShapeProperties ctLegendSpPr = ctLegend.addNewSpPr();
        ctLegendSpPr.addNewNoFill();
        ctLegendSpPr.addNewLn().addNewNoFill();
        ctLegendSpPr.addNewEffectLst();

        CTTextBody ctLegendsTxPr = ctLegend.addNewTxPr();
        CTTextBodyProperties ctLegendBodyPr = ctLegendsTxPr.addNewBodyPr();
        ctLegendBodyPr.setRot(0);
        ctLegendBodyPr.setSpcFirstLastPara(true);
        ctLegendBodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
        ctLegendBodyPr.setVert(STTextVerticalType.HORZ);
        ctLegendBodyPr.setWrap(STTextWrappingType.SQUARE);
        ctLegendBodyPr.setAnchor(STTextAnchoringType.CTR);
        ctLegendBodyPr.setAnchorCtr(true);
        ctLegendsTxPr.addNewLstStyle();
        
        CTTextParagraph ctLegendP = ctLegendsTxPr.addNewP();
        
        CTTextCharacterProperties ctLegendDefRPr = ctLegendP.addNewPPr().addNewDefRPr();
        ctLegendDefRPr.setSz(900);
        ctLegendDefRPr.setKern(1200);
        ctLegendDefRPr.setI(false);
        ctLegendDefRPr.setB(false);
        ctLegendDefRPr.setU(STTextUnderlineType.NONE);
        ctLegendDefRPr.setStrike(STTextStrikeType.NO_STRIKE);
        ctLegendDefRPr.setBaseline(0);
        
        CTSchemeColor ctLegendPSchemeClr = ctLegendDefRPr.addNewSolidFill().addNewSchemeClr();
        ctLegendPSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctLegendPSchemeClr.addNewLumMod().setVal(65000);
        ctLegendPSchemeClr.addNewLumOff().setVal(35000);
        
        ctLegendDefRPr.addNewLatin().setTypeface("+mn-lt");
        ctLegendDefRPr.addNewEa().setTypeface("+mn-ea");
        ctLegendDefRPr.addNewCs().setTypeface("+mn-cs");
        
        ctLegendP.addNewEndParaRPr().setLang("en-US");
        
        ctChart.addNewPlotVisOnly().setVal(true	);
        ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.GAP);
        ctChart.addNewShowDLblsOverMax().setVal(false);
        
        CTShapeProperties ctSpaceSpPr = ctChartSpace.addNewSpPr();
        ctSpaceSpPr.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.BG_1);
        CTLineProperties catSpaceLn = ctSpaceSpPr.addNewLn();
        catSpaceLn.setW(9525);
        catSpaceLn.setCap(STLineCap.FLAT);
        catSpaceLn.setCmpd(STCompoundLine.SNG);
        catSpaceLn.setAlgn(STPenAlignment.CTR);
        
        CTSchemeColor ctSpaceSchemeClr = catSpaceLn.addNewSolidFill().addNewSchemeClr();
        ctSpaceSchemeClr.setVal(STSchemeColorVal.TX_1);
        ctSpaceSchemeClr.addNewLumMod().setVal(15000);
        ctSpaceSchemeClr.addNewLumOff().setVal(85000);
        
        catSpaceLn.addNewRound();
        
        ctSpaceSpPr.addNewEffectLst();
        
        CTTextBody ctSpaceTxPr = ctChartSpace.addNewTxPr();
        ctSpaceTxPr.addNewBodyPr();
        ctSpaceTxPr.addNewLstStyle();
        
        CTTextParagraph ctSpaceP = ctSpaceTxPr.addNewP();
        ctSpaceP.addNewPPr().addNewDefRPr();
        ctSpaceP.addNewEndParaRPr().setLang("en-US");
        ctChartSpace.getExternalData().addNewAutoUpdate().setVal(false);
        System.out.println(chart);
	}
}
