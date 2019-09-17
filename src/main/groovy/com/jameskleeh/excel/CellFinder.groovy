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
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.streaming.SXSSFCell

/**
 * A class to get references to cells for use in other functions
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class CellFinder {

    private final SXSSFCell cell
    private final Map<Object, Integer> columnIndexes

    CellFinder(SXSSFCell cell, Map<Object, Integer> columnIndexes) {
        this.cell = cell
        this.columnIndexes = columnIndexes
    }

    /**
     * @return The current row number (1 based)
     */
    int getRow() {
        cell.rowIndex + 1
    }

    /**
     * @return The current column (A..Z)
     */
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

    /**
     * Retrieves a cell relative to the current cell. Use negative values to reference rows and columns before the current cell and positive values to reference rows and columns after the current cell.
     *
     * @param columnIndex The column index
     * @param rowIndex The row index
     * @return A cell relative to the current cell
     */
    String relativeCell(int columnIndex, int rowIndex) {
        relativeColumn(columnIndex) + relativeRow(rowIndex)
    }

    /**
     * Retrieves a cell relative to the current cell. Use negative values to reference previous columns and positive values to reference columns after the current cell.
     *
     * @param columnIndex The column index
     * @return A cell relative to the current cell
     */
    String relativeCell(int columnIndex) {
        relativeCell(columnIndex, 0)
    }

    /**
     * Retrieves an exact cell reference
     *
     * @param columnIndex The column index
     * @param rowIndex The row index
     * @return The desired cell
     */
    String exactCell(int columnIndex, int rowIndex) {
        if (rowIndex < 0) {
            throw new IllegalArgumentException("An invalid row index of $rowIndex was specified")
        }
        exactColumn(columnIndex) + (rowIndex + 1)
    }

    /**
     * Retrieves an exact cell reference based on a previously defined column definition (created by {@link com.jameskleeh.excel.internal.CreatesCells#column} and row index
     *
     * @param columnName The column identifier
     * @param rowIndex The row index
     * @return The desired cell
     */
    String exactCell(String columnName, int rowIndex) {
        if (columnIndexes && columnIndexes.containsKey(columnName)) {
            exactCell(columnIndexes[columnName], rowIndex)
        } else {
            throw new IllegalArgumentException("Column index not specified for $columnName")
        }
    }

    /**
     * Retrieves an exact cell reference based on a previously defined column definition (created by {@link com.jameskleeh.excel.internal.CreatesCells#column}
     *
     * @param columnName The column identifier
     * @return The actual header column
     */
    String exactCell(String columnName) {
        exactCell(columnName, 0)
    }

    /**
     * @return The current sheet name
     */
    String getSheetName() {
        cell.sheet.sheetName
    }
}
