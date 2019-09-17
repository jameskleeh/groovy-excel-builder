package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Specification

class SheetSpec extends Specification {

    void "test skipRows"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row()
                skipRows(2)
                row()
            }
        }

        when:
        Iterator<Row> rows = workbook.getSheetAt(0).rowIterator()

        then:
        rows.next().rowNum == 0
        rows.next().rowNum == 3
        !rows.hasNext()
    }

    void "test row(Object...)"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row(1, 2, 3)
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.physicalNumberOfCells == 3
        row.getCell(0).numericCellValue == 1
        row.getCell(1).numericCellValue == 2
        row.getCell(2).numericCellValue == 3
    }

    void "test row(Map options)"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row([height: 12F]) {

                }
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.heightInPoints == 12F
    }
}
