package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A class used to create a row in an excel document
 */
@CompileStatic
class Row {

    private final XSSFRow row
    private final XSSFWorkbook workbook
    private final XSSFSheet sheet
    private Map defaultOptions
    private final Map<Object, Integer> columnIndexes
    private final CellStyleBuilder styleBuilder
    private int cellIdx

    Row(XSSFWorkbook workbook, XSSFSheet sheet, XSSFRow row, Map defaultOptions, Map<Object, Integer> columnIndexes) {
        this.workbook = workbook
        this.sheet = sheet
        this.row = row
        this.cellIdx = 0
        this.defaultOptions = defaultOptions
        this.columnIndexes = columnIndexes
        this.styleBuilder = new CellStyleBuilder(workbook)
    }

    private XSSFCell nextCell() {
        XSSFCell cell = row.createCell(cellIdx)
        cellIdx++
        cell
    }

    private void setStyle(Object value, XSSFCell cell, Map options) {
        styleBuilder.setStyle(value, cell, options, defaultOptions)
    }

    void skipCells(int num) {
        cellIdx += num
    }

    void skipTo(Object id) {
        if (columnIndexes && columnIndexes.containsKey(id)) {
            cellIdx = columnIndexes[id]
        } else {
            throw new IllegalArgumentException("Column index not specified for $id")
        }
    }

    void defaultStyle(Map options) {
        this.defaultOptions = options
    }

    XSSFCell column(String value, Object id, final Map options = [:]) {
        XSSFCell cell = nextCell()
        cell.setCellValue(value)
        setStyle(value, cell, options)
        columnIndexes[id] = cell.columnIndex
        cell
    }

    XSSFCell formula(String formulaString, final Map style) {
        XSSFCell cell = nextCell()
        if (formulaString.startsWith('=')) {
            formulaString = formulaString[1..-1]
        }
        cell.setCellFormula(formulaString)
        setStyle(null, cell, style)
        cell
    }

    XSSFCell formula(String formulaString) {
        formula(formulaString, null)
    }

    XSSFCell formula(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Formula) Closure callable) {
        formula(null, callable)
    }

    XSSFCell formula(final Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Formula) Closure callable) {
        XSSFCell cell = nextCell()
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new Formula(cell, columnIndexes)
        String formula
        if (callable.maximumNumberOfParameters == 1) {
            formula = (String)callable.call(cell)
        } else {
            formula = (String)callable.call()
        }
        if (formula.startsWith('=')) {
            formula = formula[1..-1]
        }
        cell.setCellFormula(formula)
        setStyle(null, cell, style)
        cell
    }

    XSSFCell cell() {
        XSSFCell cell = nextCell()
        cell.setCellValue('')
        cell
    }
    XSSFCell cell(Object value) {
        cell(value, null)
    }
    XSSFCell cell(Object value, final Map style) {

        XSSFCell cell = nextCell()
        setStyle(value, cell, style)
        if (value instanceof String) {
            cell.setCellValue(value)
        } else if (value instanceof Calendar) {
            cell.setCellValue(value)
        } else if (value instanceof Date) {
            cell.setCellValue(value)
        } else if (value instanceof Number) {
            cell.setCellValue(value.doubleValue())
        } else if (value instanceof Boolean) {
            cell.setCellValue(value)
        } else {
            Closure callable = Excel.getRenderer(value.class)
            if (callable != null) {
                cell.setCellValue((String)callable.call(value))
            } else {
                cell.setCellValue(value.toString())
            }
        }
        cell
    }

}
