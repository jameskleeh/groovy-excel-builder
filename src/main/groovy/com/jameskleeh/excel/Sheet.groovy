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
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A class used to create a sheet in an excel document
 */
@CompileStatic
class Sheet {

    private final XSSFSheet sheet
    private final XSSFWorkbook workbook
    private int rowIdx
    private int columnIdx
    private Map defaultOptions
    private Map<Object, Integer> columnIndexes = [:]
    private final CellStyleBuilder styleBuilder

    private static final String HEIGHT = 'height'

    Sheet(XSSFWorkbook workbook, XSSFSheet sheet, CellStyleBuilder styleBuilder) {
        this.workbook = workbook
        this.sheet = sheet
        this.rowIdx = 0
        this.columnIdx = 0
        this.styleBuilder = styleBuilder
    }

    void defaultStyle(Map options) {
        options = new LinkedHashMap(options)
        styleBuilder.convertSimpleOptions(options)
        defaultOptions = options
    }

    void skipRows(int num) {
        rowIdx += num
    }

    void columns(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(callable)
    }

    void columns(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(options, callable)
    }

    void column(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new Column(workbook, sheet, defaultOptions, columnIndexes, styleBuilder, columnIdx, rowIdx)
        callable.call()
        columnIdx++
    }

    XSSFRow row() {
        row([:], null)
    }

    XSSFRow row(Object...cells) {
        row {
            cells.each { val ->
                cell(val)
            }
        }
    }

    XSSFRow row(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row([:], callable)
    }

    XSSFRow row(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        XSSFRow row = sheet.createRow(rowIdx)
        if (options?.containsKey(HEIGHT)) {
            Object height = options[HEIGHT]
            if (height instanceof Short) {
                row.setHeight(height)
            } else if (height instanceof Float) {
                row.setHeightInPoints(height)
            } else {
                throw new IllegalArgumentException('Row height must be a short or float')
            }
        }

        if (callable != null) {
            callable.resolveStrategy = Closure.DELEGATE_FIRST
            callable.delegate = new Row(workbook, sheet, row, defaultOptions, columnIndexes, styleBuilder)
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
