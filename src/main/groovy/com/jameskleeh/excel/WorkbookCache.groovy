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
