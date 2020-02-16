package com.himanshu.excel.parser.sax;

import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Maps;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Map;

public class SheetHeaderDataHandler extends DefaultHandler {
  private Logger logger = LoggerFactory.getLogger(SheetHeaderDataHandler.class);
  private Map<String, String> headerMap = Maps.newHashMap();
  private String content;
  private int currentRowCtr, currentRow;
  private String cellType, currentCell;
  private final SharedStringsTable sst;

  public SheetHeaderDataHandler(SharedStringsTable sst) {
    this.sst = sst;
  }

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
    content = "";
    if (currentRowCtr <= 1) {
      switch (qName) {
        case "row":
          currentRowCtr++;
          currentRow = Integer.parseInt(attributes.getValue("r"));
          break;
        case "c":
          String cellId = attributes.getValue("r");
          currentCell = cellId.substring(0, cellId.lastIndexOf(String.valueOf(currentRow)));
          cellType = attributes.getValue("t");
          break;
        case "v":
          break;
        default:
          break;
      }
    }
  }

  @Override
  public void endElement(String uri, String localName, String qName) throws SAXException {
    if(currentRowCtr > 1) {
      return;
    }
    if (qName.equalsIgnoreCase("c")) {
      if (cellType.equalsIgnoreCase("s")) {
        content = sst.getItemAt(Integer.parseInt(content)).getString();
      } else {
        //NOOP, no need to lookup from sst
      }
      logger.info("Header val at cell: [{}] is [{}]", currentCell, content);
      headerMap.put(currentCell, content);
    }

  }

  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    if(currentRowCtr > 1) {
      return;
    }
    content += new String(ch, start, length);
  }

  public Map<String, String> getHeaderMap() {
    return ImmutableMap.copyOf(headerMap);
  }
}