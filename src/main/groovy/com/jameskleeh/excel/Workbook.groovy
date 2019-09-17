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
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook

/**
 * A class used to create a workbook in an excel document
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class Workbook {

    private final SXSSFWorkbook wb
    private final CellStyleBuilder styleBuilder

    private static final String WIDTH = 'width'
    private static final String HEIGHT = 'height'

    Workbook(SXSSFWorkbook wb) {
        this.wb = wb
        this.styleBuilder = new CellStyleBuilder(wb)
    }

    /**
     * Creates a sheet
     *
     * @param name The sheet name
     * @param options Default sheet options
     * @param callable To build data
     * @return The native sheet object
     */
    SXSSFSheet sheet(String name, Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(WorkbookUtil.createSafeSheetName(name)), options, callable)
    }

    /**
     * Creates a sheet
     *
     * @param name The sheet name
     * @param callable To build data
     * @return The native sheet object
     */
    SXSSFSheet sheet(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet(name, [:], callable)
    }

    /**
     * Creates a sheet
     *
     * @param callable To build data
     * @return The native sheet object
     */
    SXSSFSheet sheet(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet([:], callable)
    }

    /**
     * Creates a sheet
     *
     * @param options Default sheet options
     * @param callable To build data
     * @return The native sheet object
     */
    SXSSFSheet sheet(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(), options, callable)
    }

    private SXSSFSheet handleSheet(SXSSFSheet sheet, Map options, Closure callable) {
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        if (options.containsKey(WIDTH)) {
            Object width = options[WIDTH]
            if (width instanceof Integer) {
                sheet.setDefaultColumnWidth(width)
            } else {
                throw new IllegalArgumentException('Sheet default column width must be an integer')
            }
        }

        if (options.containsKey(HEIGHT)) {
            Object height = options[HEIGHT]
            if (height instanceof Short) {
                sheet.setDefaultRowHeight(height)
            } else if (height instanceof Float) {
                sheet.setDefaultRowHeightInPoints(height)
            } else {
                throw new IllegalArgumentException('Sheet default row height must be a short or float')
            }
        }

        callable.delegate = new Sheet(sheet, styleBuilder)
        if (callable.maximumNumberOfParameters == 1) {
            callable.call(sheet)
        } else {
            callable.call()
        }
        sheet
    }
}
