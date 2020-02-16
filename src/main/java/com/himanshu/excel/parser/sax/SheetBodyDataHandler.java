package com.himanshu.excel.parser.sax;

import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Maps;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.List;
import java.util.Map;

public class SheetBodyDataHandler extends DefaultHandler {
  private Logger logger = LoggerFactory.getLogger(SheetBodyDataHandler.class);
  private String content;
  private int currentRowCtr, currentRow;
  private String cellType, currentCell;
  private final SharedStringsTable sst;
  private final Map<String, String> headerMap;
  private List<Map<String, String>> bodyMap;
  private Map<String, String> currentRowBody;

  public SheetBodyDataHandler(SharedStringsTable sst, Map<String, String> headerMap) {
    this.sst = sst;
    this.headerMap = headerMap;
    this.bodyMap = Lists.newArrayList();
  }

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
    content = "";
    if (currentRowCtr >= 1) {
      switch (qName) {
        case "row":
          //currentRowCtr++;
          currentRow = Integer.parseInt(attributes.getValue("r"));
          logger.info("Parsing row: {}", currentRow);
          currentRowBody = Maps.newHashMap();
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
    if (currentRowCtr < 1 && qName.equalsIgnoreCase("row")) {
      currentRowCtr++; //By-pass header row
      return;
    }
    if(currentRow == 0) {
      return;
    }
    if (qName.equalsIgnoreCase("row")) {
      bodyMap.add(currentRowBody);
    }
    if (qName.equalsIgnoreCase("c")) {
      if (cellType != null && cellType.equalsIgnoreCase("s")) {
        content = sst.getItemAt(Integer.parseInt(content)).getString();
      } else {
        //NOOP, no need to lookup from sst
      }
      logger.info("Body val at cell: [{}] is [{}]", currentCell, content);
      currentRowBody.put(headerMap.get(currentCell), content);
    }

  }

  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    if(currentRowCtr < 1) {
      return;
    }
    content += new String(ch, start, length);
  }

  public List<Map<String, String>> getBodyMap() {
    return ImmutableList.copyOf(bodyMap);
  }
}