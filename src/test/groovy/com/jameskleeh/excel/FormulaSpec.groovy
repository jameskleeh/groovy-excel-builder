package com.jameskleeh.excel

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification
import spock.lang.Subject

@Subject(CellFinder)
class FormulaSpec extends Specification {

    void "test getRow"() {
        given:
        XSSFCell cell = new XSSFWorkbook().createSheet().createRow(2).createCell(0)

        CellFinder formula = new CellFinder(cell, null)

        expect:
        formula.row == 3
    }

    void "test getColumn"() {
        given:
        XSSFCell cell = new XSSFWorkbook().createSheet().createRow(2).createCell(2)

        CellFinder formula = new CellFinder(cell, null)

        expect:
        formula.column == 'C'
    }

    void "test relativeCell(int columnIndex, int rowIndex)"() {
        given:
        XSSFCell cell = new XSSFWorkbook().createSheet().createRow(2).createCell(2)
        CellFinder formula = new CellFinder(cell, null)

        when:
        formula.relativeCell(-3, 0)

        then:
        thrown(IllegalArgumentException)

        when:
        formula.relativeCell(0, -3)

        then:
        thrown(IllegalArgumentException)

        when:
        String relativeCell = formula.relativeCell(column, row)

        then:
        relativeCell == result

        where:
        column  | row   | result
        0       | 0     | 'C3'
        -1      | 0     | 'B3'
        0       | -1    | 'C2'
        -1      | -1    | 'B2'
        1       | 0     | 'D3'
        0       | 1     | 'C4'
        1       | 1     | 'D4'
    }

    void "test relativeCell(int columnIndex)"() {
        given:
        XSSFCell cell = new XSSFWorkbook().createSheet().createRow(2).createCell(2)
        CellFinder formula = new CellFinder(cell, null)

        when:
        formula.relativeCell(-3)

        then:
        thrown(IllegalArgumentException)

        when:
        String relativeCell = formula.relativeCell(column)

        then:
        relativeCell == result

        where:
        column  | result
        0       | 'C3'
        -1      | 'B3'
        1       | 'D3'
    }

    void "test exactCell(int columnIndex, int rowIndex)"() {
        given:
        CellFinder formula = new CellFinder(null, null)

        when:
        formula.exactCell(-1, 0)

        then:
        thrown(IllegalArgumentException)

        when:
        formula.exactCell(0, -1)

        then:
        thrown(IllegalArgumentException)

        when:
        String cell = formula.exactCell(column, row)

        then:
        cell == result

        where:
        column  | row   | result
        0       | 0     | 'A1'
        1       | 0     | 'B1'
        2       | 0     | 'C1'
        0       | 1     | 'A2'
        0       | 2     | 'A3'
    }

    void "test exactCell based on column name"() {
        given:
        CellFinder formula = new CellFinder(null, ['foo': 0, 'bar': 2])

        when:
        formula.exactCell('x', 0)

        then:
        thrown(IllegalArgumentException)

        when:
        formula.exactCell('foo', -1)

        then:
        thrown(IllegalArgumentException)

        when:
        formula.exactCell('x')

        then:
        thrown(IllegalArgumentException)

        expect:
        formula.exactCell('foo') == 'A1'
        formula.exactCell('bar') == 'C1'

        when:
        String cell = formula.exactCell(column, row)

        then:
        cell == result

        where:
        column  | row   | result
        'foo'   | 0     | 'A1'
        'foo'   | 1     | 'A2'
        'bar'   | 0     | 'C1'
        'bar'   | 2     | 'C3'
    }

    void "test getSheetName"() {
        given:
        XSSFCell cell = new XSSFWorkbook().createSheet('Foo').createRow(2).createCell(2)
        CellFinder formula = new CellFinder(cell, null)

        expect:
        formula.sheetName == 'Foo'
    }

}
