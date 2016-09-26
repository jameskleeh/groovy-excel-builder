package com.jameskleeh.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.Row
import spock.lang.Specification

class SheetSpec extends Specification {

    void "test skipRows"() {
        XSSFWorkbook workbook = new ExcelBuilder().build {
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
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                row(1,2,3)
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.getPhysicalNumberOfCells() == 3
        row.getCell(0).stringCellValue == '1'
        row.getCell(1).stringCellValue == '2'
        row.getCell(2).stringCellValue == '3'
    }

    void "test row(Map options)"() {
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                row([height: 12]) {

                }
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.heightInPoints == 12F
    }
}
