package com.himanshu.excel.parser.sax;

import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Maps;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Map;

public class WorkbookDataHandler extends DefaultHandler {
  private Logger logger = LoggerFactory.getLogger(WorkbookDataHandler.class);
  private Map<String, String> sheetNameToRelIdMap = Maps.newHashMap();
  private String content;

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
    content = null;
    if(qName.equalsIgnoreCase("sheet")) {
      String sheetName = attributes.getValue("name");
      String sheetId = attributes.getValue("r:id");
      sheetNameToRelIdMap.put(sheetName, sheetId);
      logger.info("Found sheet start");
    }
  }

  @Override
  public void endElement(String uri, String localName, String qName) throws SAXException {
    if(qName.equalsIgnoreCase("sheet")) {
      logger.info("Found sheet end");
    }
  }

  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    content += new String(ch, start, length);
  }

  public Map<String, String> getSheetNameToRelIdMap() {
    return ImmutableMap.copyOf(sheetNameToRelIdMap);
  }
}
