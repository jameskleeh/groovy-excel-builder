package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell

/**
 * A class to get references to cells for use in formulas
 *
 * @author James Kleeh
 * @since 1.0.0
 */
@CompileStatic
class Formula {

    private final XSSFCell cell
    private final Map<Object, Integer> columnIndexes

    Formula(XSSFCell cell, Map<Object, Integer> columnIndexes) {
        this.cell = cell
        this.columnIndexes = columnIndexes
    }

    int getRow() {
        cell.rowIndex + 1
    }

    String getColumn() {
        relativeColumn(0)
    }

    private int relativeRow(int index) {
        int rowIndex = row + index
        if (rowIndex < 1) {
            throw new IllegalArgumentException("An invalid row index of $rowIndex was specified")
        }
        rowIndex
    }

    private String relativeColumn(int index) {
        exactColumn(cell.columnIndex + index)
    }

    private String exactColumn(int index) {
        if (index > -1) {
            CellReference.convertNumToColString(index)
        } else {
            throw new IllegalArgumentException("An invalid column index of $index was specified")
        }
    }

    String relativeCell(int columnIndex, int rowIndex) {
        relativeColumn(columnIndex) + relativeRow(rowIndex)
    }

    String relativeCell(int columnIndex) {
        relativeCell(columnIndex, 0)
    }

    String exactCell(int columnIndex, int rowIndex) {
        if (rowIndex < 0) {
            throw new IllegalArgumentException("An invalid row index of $rowIndex was specified")
        }
        exactColumn(columnIndex) + (rowIndex + 1)
    }

    String exactCell(String columnName, int rowIndex) {
        if (columnIndexes && columnIndexes.containsKey(columnName)) {
            exactCell(columnIndexes[columnName], rowIndex)
        } else {
            throw new IllegalArgumentException("Column index not specified for $columnName")
        }
    }

    String exactCell(String columnName) {
        exactCell(columnName, 0)
    }
}
