package com.himanshu.excel.parser.sax;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class SAXExcelReaderTest {
  private static final Logger logger = LoggerFactory.getLogger(SAXExcelReaderTest.class);

  @Test
  @DisplayName("Test if file exists")
  public void testParsing() {
    String filePath = SAXExcelReaderTest.class.getResource("/").getFile().concat("/sample.xlsx");
    File f = new File(filePath);
    Assertions.assertAll(()-> Assertions.assertTrue(f.isFile()));
  }

  @Test
  @DisplayName("Parse sheet names")
  public void parseSheetNamesTest() throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
    String filePath = SAXExcelReaderTest.class.getResource("/").getFile().concat("/sample.xlsx");
    File f = new File(filePath);
    SAXExcelReader saxExcelReader = new SAXExcelReader(f, "NameVsRank");
    List<Map<String, String>> body = saxExcelReader.read();
    logger.info("{}", body);
  }
}

