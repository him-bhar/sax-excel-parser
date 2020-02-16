package com.himanshu.excel.parser.sax;

import java.util.Map;

public interface MapToObjectTransformer<T> {
  T transform(Map<String, String> values);
}
