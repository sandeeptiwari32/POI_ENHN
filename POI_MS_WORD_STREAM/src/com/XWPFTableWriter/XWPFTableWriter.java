package com.XWPFTableWriter;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.util.TempFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
/**
 * 
 * @author Sandeep Tiwari
 *This class is used to write XWPFTable xml in temp file and then write it to doc file
 */
public class XWPFTableWriter {
	private static final String MAIN="main:";
	private static final String W="w:";
	private static final String TABLE_START_TAG="<w:tbl>";
	private static final String TABLE_END_TAG="</w:tbl>";
	private static final String TABLE_IDENTIFICATION_TAG="<w:tbl";
	private static final String TABLE_PR_TAG = "<w:tblPr>";
	private static final String XML_END_TAG = "</xml-fragment>";
	private static final String DOCUMENT_XML = "word/document.xml";
	private static final String POI_DOC_TEMP_FILE_NAME = "POI_DOCUMENT_FILE";
	private static final String DOC_FILE_EXT = ".docx";
	private static final String XML_FILE_EXT = ".xml";
	private ArrayList<String> tempFilePath=new ArrayList<String>();
	private XWPFDocument document=null;

	/** 
	 * @param document
	 * @throws IOException
	 */
	public XWPFTableWriter(XWPFDocument document) throws IOException {
		this.document=document;
	}
	/**
	 * 
	 * @param table
	 * @return byte string of XWPFTable table
	 */
	private byte[] getTableBytes(XWPFTable table) {
		try {
			String xml = table.getCTTbl().toString();
			//replace name space from main to w
			xml = xml.replace(MAIN,W);
			//inclose in table tag
			xml = TABLE_START_TAG +xml.substring(xml.indexOf(TABLE_PR_TAG), xml.indexOf(XML_END_TAG))+TABLE_END_TAG;
			//return xml in byte
			return xml.getBytes("UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * inject data in output stream from temp files 
	 * @param zipfile
	 * @param out
	 * @throws IOException
	 */
	private void injectData(File zipfile, OutputStream out) throws IOException {
		ZipFile zip = ZipHelper.openZipFile(zipfile);
		try {
			ZipOutputStream zos = new ZipOutputStream(out);
			try {
				Enumeration<? extends ZipEntry> en = zip.entries();
				while (en.hasMoreElements()) {

					ZipEntry ze = en.nextElement();
					zos.putNextEntry(new ZipEntry(ze.getName()));
					InputStream is = zip.getInputStream(ze);
					if (ze.getName().equalsIgnoreCase(DOCUMENT_XML)) {
						//copy xml data of temp file in output stream
						copyStreamAndInjectInDocument(is, zos);
					} else {
						//copy rest content in output stream
						copyStream(is, zos);
					}
					is.close();
				}
			} finally {
				zos.close();
			}
		} finally {
			zip.close();
		}
	}

	/**
	 * copy rest data
	 * @param in
	 * @param out
	 * @throws IOException
	 */
	private void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] chunk = new byte[1024];
		int count;
		while ((count = in.read(chunk)) >= 0) {
			out.write(chunk, 0, count);
		}
	}

	/**
	 * copy data
	 * @param in
	 * @param out
	 * @throws IOException
	 */
	private void copyStreamAndInjectInDocument(InputStream in, OutputStream out)
			throws IOException {
		BufferedReader inReader = new BufferedReader(new InputStreamReader(in, "UTF-8"));
		OutputStreamWriter outWriter = new OutputStreamWriter(out, "UTF-8");
		// Copy from "in" to "out" up to the string "<w:tbl/>" or
		// "</w:tbl>" (excluding).
			String tempLine="";
			while ((tempLine = inReader.readLine())!= null) {
				//write data in file till we reach table tag
				//we can change it as per our requirement
				if (tempLine.contains(TABLE_IDENTIFICATION_TAG)) {
					int index=0;
					while(tempLine.indexOf(TABLE_IDENTIFICATION_TAG)!=-1)
					{
						tempLine=this.writeXMLData(outWriter,tempLine,index);
						index++;
					}
					outWriter.write(tempLine);
				} else {
					outWriter.write(tempLine);
				}
			}
			outWriter.flush();
	}
	/**
	 * write xml data into doc file
	 * @param outWriter
	 * @param tempLine
	 * @param index
	 * @return substring string
	 * @throws IOException
	 */
	private String writeXMLData(OutputStreamWriter outWriter, String tempLine, int index) throws IOException {
		outWriter.write(tempLine.substring(0, tempLine.indexOf(TABLE_IDENTIFICATION_TAG)));
		File tempFile=new File(tempFilePath.get(index));
		BufferedReader documentData = new BufferedReader(new InputStreamReader(new FileInputStream(tempFile), "UTF-8"));
		copyStream(documentData, outWriter);
		documentData.close();
		if (!tempFile.delete()) {
			throw new IOException("Could not delete temporary file after processing: " + tempFile);
		}
		return tempLine.substring(tempLine.indexOf(TABLE_END_TAG)+TABLE_END_TAG.length());
	}
	/**
	 * write temp xml file stream into output stream
	 * @param documentData
	 * @param outWriter
	 * @throws IOException
	 */
	private void copyStream(BufferedReader documentData, OutputStreamWriter outWriter) throws IOException {
		String tempLine="";
		while ((tempLine = documentData.readLine())!= null) {
			outWriter.write(tempLine);
		}
	}
	/**
	 * write XWPFDocument into output stream
	 * @param stream
	 * @throws IOException
	 */
	public void writeDocument(OutputStream stream) throws IOException {
		// create temp doc file
		File tmplFile = TempFile.createTempFile(POI_DOC_TEMP_FILE_NAME, DOC_FILE_EXT);
		try {
			FileOutputStream os = new FileOutputStream(tmplFile);
			try {
				this.document.write(os);
			} finally {
				os.close();
			}
			// Substitute the template entries with the generated document xml data
			// files
			this.injectData(tmplFile, stream);
		} finally {
			if (!tmplFile.delete()) {
				throw new IOException("Could not delete temporary file after processing: " + tmplFile);
			}
		}
	}
	/**
	 * write table xml in document xml file temp file
	 * @param tableOne
	 * @throws IOException
	 */
	public void writeTable(XWPFTable table) throws IOException {
		//create temp xml file to store rows of XWPFTable
		File file = TempFile.createTempFile(POI_DOC_TEMP_FILE_NAME,XML_FILE_EXT);
		FileOutputStream fos = new FileOutputStream(file);
		fos.write(getTableBytes(table));
		fos.flush();		
		fos.close();
		tempFilePath.add(file.getAbsolutePath());
		//clear table
		//we have data in temp xml file
		table.getCTTbl().getTrList().clear();
	}
}
