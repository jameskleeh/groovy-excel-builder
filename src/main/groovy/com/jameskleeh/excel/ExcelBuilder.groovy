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
import org.apache.poi.xssf.streaming.SXSSFWorkbook

/**
 * The main class used to start building an excel document
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class ExcelBuilder {

    /**
     * Builds an excel document and sends the data to an output stream. The output stream is NOT closed.
     *
     * @param outputStream An output stream to push data onto
     * @param callable The closure to build the document
     */
    static void output(OutputStream outputStream, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Workbook) Closure callable) {
        SXSSFWorkbook wb = build(callable)
        wb.write(outputStream)
    }

    /**
     * Builds an excel document
     *
     * @param callable The closure to build the document
     * @return The native workbook
     */
    static SXSSFWorkbook build(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Workbook) Closure callable) {
        SXSSFWorkbook wb = new SXSSFWorkbook()
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new Workbook(wb)
        if (callable.maximumNumberOfParameters == 1) {
            callable.call(wb)
        } else {
            callable.call()
        }
        wb
    }
}
