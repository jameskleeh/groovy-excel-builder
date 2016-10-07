/*
Licensed to the Apache Software Foundation (ASF) under one
or more contributor license agreements.  See the NOTICE file
distributed with this work for additional information
regarding copyright ownership.  The ASF licenses this file
to you under the Apache License, Version 2.0 (the
        "License"); you may not use this file except in compliance
with the License.  You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing,
        software distributed under the License is distributed on an
"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, either express or implied.  See the License for the
specific language governing permissions and limitations
under the License.
*/
package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A class used to store fonts and styles for reuse in workbooks
 */
@CompileStatic
class WorkbookCache {

    final Map<Object, XSSFFont> fonts = [:]
    final Map<Object, XSSFCellStyle> styles = [:]

    private final XSSFWorkbook workbook

    WorkbookCache(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    Boolean containsFont(Object obj) {
        fonts.containsKey(obj)
    }

    Boolean containsStyle(Object obj) {
        styles.containsKey(obj)
    }

    XSSFFont getFont(Object obj) {
        fonts.get(obj)
    }

    XSSFCellStyle getStyle(Object obj) {
        styles.get(obj)
    }

    void putFont(Object obj, XSSFFont font) {
        fonts.put(obj, font)
    }

    void putStyle(Object obj, XSSFCellStyle style) {
        styles.put(obj, style)
    }
}
