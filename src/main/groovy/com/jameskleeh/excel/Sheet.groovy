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
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet

/**
 * A class used to create a sheet in an excel document
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class Sheet {

    private final SXSSFSheet sheet
    private int rowIdx
    private int columnIdx
    private Map defaultOptions
    private Map<Object, Integer> columnIndexes = [:]
    private final CellStyleBuilder styleBuilder

    private static final String HEIGHT = 'height'
    private static final String WIDTH = 'width'

    Sheet(SXSSFSheet sheet, CellStyleBuilder styleBuilder) {
        this.sheet = sheet
        this.rowIdx = 0
        this.columnIdx = 0
        this.styleBuilder = styleBuilder
    }

    /**
     * Sets the default styling for the sheet
     *
     * @param options Style options
     */
    void defaultStyle(Map options) {
        options = new LinkedHashMap(options)
        styleBuilder.convertSimpleOptions(options)
        defaultOptions = options
    }

    /**
     * Skips rows
     * @param num The number of rows to skip
     */
    void skipRows(int num) {
        rowIdx += num
    }

    /**
     * Skips columns
     * @param num The number of columns to skip
     */
    void skipColumns(int num) {
        columnIdx += num
    }

    /**
     * Used to define headers for a sheet
     *
     * @param callable To build header data
     */
    void columns(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(callable)
    }

    /**
     * Used to define headers for a sheet
     *
     * @param options Default style options for the header
     * @param callable To build header data
     */
    void columns(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(options, callable)
    }

    /**
     * Output data by column
     *
     * @param callable To build column data
     */
    void column(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new Column(sheet, defaultOptions, columnIndexes, styleBuilder, columnIdx, rowIdx)
        callable.call()
        columnIdx++
    }

    /**
     * Output data by column
     *
     * @param callable To build column data
     */
    void column(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        if (options?.containsKey(WIDTH)) {
            Object width = options[WIDTH]
            if (width instanceof Integer) {
                sheet.setColumnWidth(columnIdx, (Integer)width)
            } else {
                throw new IllegalArgumentException('Column width must be an integer')
            }
        }
        column(callable)
    }

    /**
     * Skip to a previously defined column created by {@link #column}
     *
     * @param id The column identifier
     */
    void skipTo(Object id) {
        if (columnIndexes && columnIndexes.containsKey(id)) {
            columnIdx = columnIndexes[id]
        } else {
            throw new IllegalArgumentException("Column index not specified for $id")
        }
    }

    /**
     * Creates a row
     *
     * @return The native row
     */
    SXSSFRow row() {
        row([:], null)
    }

    /**
     * Creates a row
     *
     * @param cells A list of data to output as cells
     * @return The native row
     */
    SXSSFRow row(Object...cells) {
        row {
            cells.each { val ->
                cell(val)
            }
        }
    }

    /**
     * Creates a row
     *
     * @param callable To build row data
     * @return The native row
     */
    SXSSFRow row(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row([:], callable)
    }

    /**
     * Creates a row
     *
     * @param options Default styling options
     * @param callable To build row data
     * @return The native row
     */
    SXSSFRow row(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        SXSSFRow row = sheet.createRow(rowIdx)
        if (options?.containsKey(HEIGHT)) {
            Object height = options[HEIGHT]
            if (height instanceof Short) {
                row.setHeight((Short)height)
            } else if (height instanceof Float) {
                row.setHeightInPoints((Float)height)
            } else {
                throw new IllegalArgumentException('Row height must be a short or float')
            }
        }

        if (callable != null) {
            callable.resolveStrategy = Closure.DELEGATE_FIRST
            callable.delegate = new Row(row, defaultOptions, columnIndexes, styleBuilder)
            if (callable.maximumNumberOfParameters == 1) {
                callable.call(row)
            } else {
                callable.call()
            }
        }
        rowIdx++
        row
    }
}
