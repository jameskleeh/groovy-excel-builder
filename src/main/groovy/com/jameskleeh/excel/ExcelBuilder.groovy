package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@CompileStatic
class ExcelBuilder {

    static void create(OutputStream outputStream, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkBook) Closure callable) {
        XSSFWorkbook wb = build(callable)
        wb.write(outputStream)
    }

    static XSSFWorkbook build(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkBook) Closure callable) {
        XSSFWorkbook wb = new XSSFWorkbook()
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new WorkBook(wb)
        if (callable.maximumNumberOfParameters == 1) {
            callable.call(wb)
        } else {
            callable.call()
        }
        wb
    }
}
