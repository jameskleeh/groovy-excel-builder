package com.jameskleeh.excel

import spock.lang.Specification

/**
 * Created by jameskleeh on 9/25/16.
 */
class ExcelSpec extends Specification {

    void "test getRenderer order"() {
        given:
        Excel.rendererEntries.clear()
        Excel.registerCellRenderer(Integer) {
            it * 2
        }
        Excel.registerCellRenderer(Integer) {
            it * 3
        }

        when:
        Closure callable = Excel.getRenderer(Integer)

        then: "Renderers registered later with the same class and priority are chosen"
        callable.call(2) == 6
    }

    void "test getRenderer higher priority"() {
        Excel.rendererEntries.clear()
        Excel.registerCellRenderer(Integer, 2) {
            it * 3
        }
        Excel.registerCellRenderer(Integer, 1) {
            it * 2
        }

        when:
        Closure callable = Excel.getRenderer(Integer)

        then: "The highest priorty renderer is chosen"
        callable.call(2) == 6
    }

    void "test getFormat higher priority"() {
        Excel.formatEntries.clear()
        Excel.registerCellFormat(Integer, 2, 2)
        Excel.registerCellFormat(Integer, 1, 1)

        when:
        Object format = Excel.getFormat(Integer)

        then: "The highest priorty renderer is chosen"
        format == 2
    }

    void "test getFormat order"() {
        Excel.formatEntries.clear()
        Excel.registerCellFormat(Integer, 2)
        Excel.registerCellFormat(Integer, 1)

        when:
        Object format = Excel.getFormat(Integer)

        then: "Formats registered later with the same class and priority are chosen"
        format == 1
    }
}
