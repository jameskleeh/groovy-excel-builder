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

import com.jameskleeh.excel.internal.CreatesCells
import com.jameskleeh.excel.style.CellRangeBorderStyleApplier
import com.jameskleeh.excel.style.ColumnCellRangeBorderStyleApplier
import groovy.transform.CompileStatic
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy

/**
 * A class designed to be a delegate when a column is created
 *
 * @author James Kleeh
 * @since 0.2.0
 */
@CompileStatic
class Column extends CreatesCells {

    private int columnIdx
    private int rowIdx

    Column(XSSFSheet sheet, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder, int columnIdx, int rowIdx) {
        super(sheet, defaultOptions, columnIndexes, styleBuilder)
        this.columnIdx = columnIdx
        this.rowIdx = rowIdx
    }

    @Override
    protected XSSFCell nextCell() {
        XSSFRow row = sheet.getRow(rowIdx)
        if (row == null) {
            row = sheet.createRow(rowIdx)
        }
        XSSFCell cell = row.getCell(columnIdx, MissingCellPolicy.CREATE_NULL_AS_BLANK)
        rowIdx++
        cell
    }

    /**
     * @see CreatesCells#skipCells
     */
    void skipCells(int num) {
        rowIdx += num
    }

    /**
     * @see CreatesCells#merge(Map style, Closure callable)
     */
    @Override
    void merge(Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        performMerge(style, callable)
    }

    @Override
    protected CellRangeAddress getRange(int start, int end) {
        new CellRangeAddress(start, end, columnIdx, columnIdx)
    }

    @Override
    protected int getMergeIndex() {
        rowIdx
    }

    @Override
    protected CellRangeBorderStyleApplier getBorderStyleApplier(CellRangeAddress range, XSSFSheet sheet) {
        new ColumnCellRangeBorderStyleApplier(range, sheet)
    }

    /**
     * @see CreatesCells#merge(Closure callable)
     */
    @Override
    void merge(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        merge(null, callable)
    }
}
