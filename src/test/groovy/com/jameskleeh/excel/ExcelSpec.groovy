package com.jameskleeh.excel

import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Issue
import spock.lang.Specification

class ExcelSpec extends Specification {

    void cleanup() {
        Excel.formatEntries.clear()
        Excel.rendererEntries.clear()
    }

    void "test getRenderer order"() {
        given:
        Excel.registerCellRenderer(Integer) {
            it * 2
        }
        Excel.registerCellRenderer(Integer) {
            it * 3
        }

        when:
        Closure callable = Excel.getRenderer(Integer)

        then: 'Renderers registered later with the same class and priority are chosen'
        callable.call(2) == 6
    }

    class Foo {

    }

    class Bar extends Foo {

    }

    void "test getRenderer subclass"() {
        given:
        Excel.registerCellRenderer(Foo) {
            it
        }

        when:
        Closure callable = Excel.getRenderer(Bar)

        then: 'Renderers registered for super classes work'
        callable.call(1) instanceof Integer
    }

    void "test getRenderer higher priority"() {
        Excel.registerCellRenderer(Integer, 2) {
            it * 3
        }
        Excel.registerCellRenderer(Integer, 1) {
            it * 2
        }

        when:
        Closure callable = Excel.getRenderer(Integer)

        then: 'The highest priorty renderer is chosen'
        callable.call(2) == 6
    }

    void "test getRenderer returns null"() {
        when:
        Closure callable = Excel.getRenderer(StringBuilder)

        then:
        callable == null
    }

    void "test getFormat higher priority"() {
        Excel.registerCellFormat(Integer, 2, 2)
        Excel.registerCellFormat(Integer, 1, 1)

        when:
        Object format = Excel.getFormat(Integer)

        then: 'The highest priorty renderer is chosen'
        format == 2
    }

    void "test getFormat order"() {
        Excel.registerCellFormat(Integer, 2)
        Excel.registerCellFormat(Integer, 1)

        when:
        Object format = Excel.getFormat(Integer)

        then: 'Formats registered later with the same class and priority are chosen'
        format == 1
    }

    void "test getFormat subclass"() {
        Excel.registerCellFormat(Foo, 1)

        when:
        Object format = Excel.getFormat(Bar)

        then: 'Formats registered for super classes work for subclasses'
        format == 1
    }

    void "test getFormat returns null"() {
        when:
        Object format = Excel.getFormat(StringBuilder)

        then:
        format == null
    }

    void "test getFormat(String) returns a built in format if it exists"() {
        Excel.registerCellFormat(Foo, 'h:mm AM/PM')

        when:
        Object format = Excel.getFormat(Foo)

        then:
        format == 18
    }

    @Issue('https://github.com/jameskleeh/groovy-excel-builder/issues/7')
    void "test creating cells with registered cell formats"() {
        Excel.registerCellRenderer(String, 0) { it + 'extra' }
        Excel.registerCellFormat(BigDecimal, 0xa)

        when:
        SXSSFWorkbook wb = ExcelBuilder.build {
            sheet {
                row {
                    cell(1.32G)
                    cell('Foo')
                }
            }
        }

        then:
        wb.getSheetAt(0).getRow(0).getCell(0).cellStyle.dataFormat == (short)0xa
        wb.getSheetAt(0).getRow(0).getCell(1).stringCellValue == 'Fooextra'
    }

}
