package com.himanshu.excel.parser.sax;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

public class SAXExcelReader {
  private static final Logger logger = LoggerFactory.getLogger(SAXExcelReader.class);
  private final File file;
  private final String sheetName;

  public SAXExcelReader(File file, String sheetName) {
    this.file = file;
    this.sheetName = sheetName;
  }

  public List<Map<String, String>> read() throws OpenXML4JException, IOException {
    OPCPackage opcPackage = OPCPackage.open(file, PackageAccess.READ);
    XSSFReader xssfReader = new XSSFReader(opcPackage);
    SAXParserFactory saxParserFactory = SAXParserFactory.newInstance();
    SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
    Map<String, String> sheetNameToIdMap = getWorkbookSheetIdentifiers(saxParserFactory, xssfReader);
    Map<String, String> headerCellToNameMap = getHeaderData(saxParserFactory, xssfReader, this.sheetName, sharedStringsTable, sheetNameToIdMap);
    List<Map<String, String>> body = getBodyData(saxParserFactory, xssfReader, this.sheetName, sharedStringsTable, sheetNameToIdMap, headerCellToNameMap);
    return body;
  }

  private Map<String, String> getHeaderData(SAXParserFactory saxParserFactory, XSSFReader xssfReader,
                             String sheetName, SharedStringsTable sharedStringsTable,
                             Map<String, String> sheetNameToIdMap) {
    try {
      String sheetId = sheetNameToIdMap.get(sheetName);
      SheetHeaderDataHandler sheetHeaderDataHandler = new SheetHeaderDataHandler(sharedStringsTable);
      XMLReader sheetReader = saxParserFactory.newSAXParser().getXMLReader();
      InputStream sheetInputStream = xssfReader.getSheet(sheetId);
      sheetReader.setContentHandler(sheetHeaderDataHandler);
      sheetReader.parse(new InputSource(sheetInputStream));
      return sheetHeaderDataHandler.getHeaderMap();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  private List<Map<String, String>> getBodyData(SAXParserFactory saxParserFactory, XSSFReader xssfReader,
                                            String sheetName, SharedStringsTable sharedStringsTable,
                                            Map<String, String> sheetNameToIdMap, Map<String, String> headerMap) {
    try {
      String sheetId = sheetNameToIdMap.get(sheetName);
      SheetBodyDataHandler sheetBodyDataHandler = new SheetBodyDataHandler(sharedStringsTable, headerMap);
      XMLReader sheetReader = saxParserFactory.newSAXParser().getXMLReader();
      InputStream sheetInputStream = xssfReader.getSheet(sheetId);
      sheetReader.setContentHandler(sheetBodyDataHandler);
      sheetReader.parse(new InputSource(sheetInputStream));
      return sheetBodyDataHandler.getBodyMap();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  private Map<String, String> getWorkbookSheetIdentifiers(SAXParserFactory saxParserFactory, XSSFReader xssfReader) {
    try {
      XMLReader workbookReader = saxParserFactory.newSAXParser().getXMLReader();
      InputStream workbookDataStream = xssfReader.getWorkbookData();
      WorkbookDataHandler workbookDataHandler = new WorkbookDataHandler();
      workbookReader.setContentHandler(workbookDataHandler);
      workbookReader.parse(new InputSource(workbookDataStream));
      return workbookDataHandler.getSheetNameToRelIdMap();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
}
