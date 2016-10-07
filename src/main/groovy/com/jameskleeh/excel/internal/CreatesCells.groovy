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
package com.jameskleeh.excel.internal

import com.jameskleeh.excel.CellStyleBuilder
import com.jameskleeh.excel.Excel
import com.jameskleeh.excel.Formula
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A base class used to create cells
 */
abstract class CreatesCells {

    protected final XSSFWorkbook workbook
    protected final XSSFSheet sheet
    protected Map defaultOptions
    protected final Map<Object, Integer> columnIndexes
    protected final CellStyleBuilder styleBuilder

    CreatesCells(XSSFWorkbook workbook, XSSFSheet sheet, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder) {
        this.workbook = workbook
        this.sheet = sheet
        this.defaultOptions = defaultOptions
        this.columnIndexes = columnIndexes
        this.styleBuilder = styleBuilder
    }


    void defaultStyle(Map options) {
        options = new LinkedHashMap(options)
        styleBuilder.convertSimpleOptions(options)
        if (defaultOptions == null) {
            defaultOptions = options
        } else {
            defaultOptions = styleBuilder.merge(defaultOptions, options)
        }
    }

    protected abstract XSSFCell nextCell()

    abstract void skipCells(int num)

    protected void setStyle(Object value, XSSFCell cell, Map options) {
        styleBuilder.setStyle(value, cell, options, defaultOptions)
    }

    XSSFCell column(String value, Object id, final Map style = null) {
        XSSFCell col = cell(value, style)
        columnIndexes[id] = col.columnIndex
        col
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
        cell('')
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

    abstract void merge(final Map style, Closure callable)

    abstract void merge(Closure callable)

    void merge(Object value, Integer count, final Map style = null) {
        merge(style) {
            cell(value)
            skipCells(count)
        }
    }

}