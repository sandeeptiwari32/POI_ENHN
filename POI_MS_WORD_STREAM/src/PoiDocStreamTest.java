import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.XWPFTableWriter.XWPFTableWriter;
/**
 * 
 * @author Sandeep Tiwari
 *this is a test file to use XWPFTABLEWriter
 */
public class PoiDocStreamTest {
	public static void main(String[] args) throws IOException {
		System.out.println("creating xwpf document");
		XWPFDocument document = new XWPFDocument();
		XWPFTableWriter tableWriter=new XWPFTableWriter(document);
		FileOutputStream outStream = new FileOutputStream("test.docx");
		
		/****************table one**********************/
		System.out.println("creating table 1");
		XWPFTable tableOne = document.createTable();
		XWPFTableRow tableOneRowOne = tableOne.getRow(0);
		tableOneRowOne.getCell(0).setText("Header1");
		tableOneRowOne.addNewTableCell().setText("header2");
		XWPFTableRow tableOneRowTwo = tableOne.createRow();
		tableOneRowTwo.getCell(0).setText("Data1");
		tableOneRowTwo.getCell(1).setText("Data2");
		System.out.println("writing table 1 rows into xml file");
		tableWriter.writeTable(tableOne);
		System.out.println("creating paragraph");
		XWPFParagraph newP = document.createParagraph();
        XWPFRun newR = newP.createRun();
        newR.setText("Sandy Tiwari");

		/*****************table two*************/
        System.out.println("creating table 2");
		tableOne = document.createTable();
		tableOneRowOne = tableOne.getRow(0);
		tableOneRowOne.getCell(0).setText("Header2");
		tableOneRowOne.addNewTableCell().setText("header3");
		tableOneRowTwo = tableOne.createRow();
		tableOneRowTwo.getCell(0).setText("Data2");
		tableOneRowTwo.getCell(1).setText("Data3");
		System.out.println("writing table 2 rows into xml file");
		tableWriter.writeTable(tableOne);
		System.out.println("creating paragraph");
		newP = document.createParagraph();
        newR = newP.createRun();
        newR.setText("Sandy Tiwari part2");
        System.out.println("writing document into output stream");
		tableWriter.writeDocument(outStream);
		outStream.close();
		System.out.println("done");
	}
}
