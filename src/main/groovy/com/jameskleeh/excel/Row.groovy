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
import com.jameskleeh.excel.style.RowCellRangeBorderStyleApplier
import groovy.transform.CompileStatic
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet

/**
 * A class used to create a row in an excel document
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class Row extends CreatesCells {

    private final SXSSFRow row

    private int cellIdx

    Row(SXSSFRow row, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder) {
        super(row.sheet, defaultOptions, columnIndexes, styleBuilder)
        this.row = row
        this.cellIdx = 0
    }

    @Override
    protected SXSSFCell nextCell() {
        SXSSFCell cell = row.createCell(cellIdx)
        cellIdx++
        cell
    }

    /**
     * @see CreatesCells#skipCells
     */
    @Override
    void skipCells(int num) {
        cellIdx += num
    }

    /**
     * Skip to a previously defined column created by {@link CreatesCells#column}
     *
     * @param id The column identifier
     */
    void skipTo(Object id) {
        if (columnIndexes && columnIndexes.containsKey(id)) {
            cellIdx = columnIndexes[id]
        } else {
            throw new IllegalArgumentException("Column index not specified for $id")
        }
    }

    /**
     * @see CreatesCells#merge(Map style, Closure callable)
     */
    @Override
    void merge(final Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        performMerge(style, callable)
    }

    @Override
    protected CellRangeAddress getRange(int start, int end) {
        new CellRangeAddress(row.rowNum, row.rowNum, start, end)
    }

    @Override
    protected int getMergeIndex() {
        cellIdx
    }

    @Override
    protected CellRangeBorderStyleApplier getBorderStyleApplier(CellRangeAddress range, SXSSFSheet sheet) {
        new RowCellRangeBorderStyleApplier(range, sheet)
    }

    /**
     * @see CreatesCells#merge(Closure callable)
     */
    @Override
    void merge(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        merge(null, callable)
    }

}
