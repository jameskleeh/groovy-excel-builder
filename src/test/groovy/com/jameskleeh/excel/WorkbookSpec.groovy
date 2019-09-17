package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Specification

class WorkbookSpec extends Specification {

    void "test sheet"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {

            }
            sheet('Other') {

            }
            sheet([height: 12F, width: 20]) {

            }
        }

        when:
        Sheet other = workbook.getSheetAt(1)
        Sheet config = workbook.getSheetAt(2)

        then:
        workbook.numberOfSheets == 3
        other.sheetName == 'Other'
        config.defaultRowHeightInPoints == 12F
        config.defaultColumnWidth == 20
    }
}
