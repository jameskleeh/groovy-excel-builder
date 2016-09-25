package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A class used to create a workbook in an excel document
 */
@CompileStatic
class WorkBook {

    private final XSSFWorkbook wb

    private static final String WIDTH = 'width'
    private static final String HEIGHT = 'height'

    WorkBook(XSSFWorkbook wb) {
        this.wb = wb
    }

    void sheet(String name, Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(WorkbookUtil.createSafeSheetName(name)), options, callable)
    }

    void sheet(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet(name, [:], callable)
    }

    void sheet(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet([:], callable)
    }

    void sheet(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(), options, callable)
    }

    private handleSheet(XSSFSheet sheet, Map options, Closure callable) {
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        if (options.containsKey(WIDTH)) {
            Object width = options[WIDTH]
            if (width instanceof Integer) {
                sheet.setDefaultColumnWidth(width)
            } else {
                throw new IllegalArgumentException('Sheet default column width must be an integer')
            }
        }
        if (options.containsKey(HEIGHT)) {
            Object height = sheet[HEIGHT]
            if (height instanceof Short) {
                sheet.setDefaultRowHeight(height)
            } else if (height instanceof Float) {
                sheet.setDefaultRowHeightInPoints(height)
            } else {
                throw new IllegalArgumentException('Sheet default row height must be a short or float')
            }
        }
        callable.delegate = new Sheet(wb, sheet)
        if (callable.maximumNumberOfParameters == 1) {
            callable.call(sheet)
        } else {
            callable.call()
        }
    }
}
