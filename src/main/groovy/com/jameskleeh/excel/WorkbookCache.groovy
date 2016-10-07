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
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class WorkbookCache {

    final Map<Object, XSSFFont> fonts = [:]
    final Map<Object, XSSFCellStyle> styles = [:]

    private final XSSFWorkbook workbook

    WorkbookCache(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    /**
     * Determines if a font has been built for a given set of options
     *
     * @param obj Font options
     * @return Whether the font exists in the cache
     */
    Boolean containsFont(Object obj) {
        fonts.containsKey(obj)
    }

    /**
     * Determines if a style has been built for a given set of options
     *
     * @param obj Style options
     * @return Whether the style exists in the cache
     */
    Boolean containsStyle(Object obj) {
        styles.containsKey(obj)
    }

    /**
     * @param obj Font options
     * @return The built font object
     */
    XSSFFont getFont(Object obj) {
        fonts.get(obj)
    }

    /**
     * @param obj Style options
     * @return The built style object
     */
    XSSFCellStyle getStyle(Object obj) {
        styles.get(obj)
    }

    /**
     * Adds a font to the cache
     *
     * @param obj Font options
     * @param font Font object
     */
    void putFont(Object obj, XSSFFont font) {
        fonts.put(obj, font)
    }

    /**
     * Adds a style to the cache
     *
     * @param obj Style options
     * @param style Style object
     */
    void putStyle(Object obj, XSSFCellStyle style) {
        styles.put(obj, style)
    }
}
