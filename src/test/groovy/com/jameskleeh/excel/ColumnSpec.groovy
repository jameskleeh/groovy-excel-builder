package com.jameskleeh.excel

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

class ColumnSpec extends Specification {

    void "test output by column"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                columns {
                    column("Column A1", "id")
                }
                column {
                    cell('A2')
                    cell('A3')
                }
                column {
                    cell('B2')
                    cell('B3')
                }
            }
        }

        when:
        XSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.getRow(0).getCell(0).getStringCellValue() == "Column A1"
        sheet.getRow(0).getCell(1) == null
        sheet.getRow(1).getCell(0).getStringCellValue() == "A2"
        sheet.getRow(1).getCell(1).getStringCellValue() == "B2"
        sheet.getRow(2).getCell(0).getStringCellValue() == "A3"
        sheet.getRow(2).getCell(1).getStringCellValue() == "B3"

    }

    void "test merge"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                column {

                }
                column {
                    merge {
                        cell('A1')
                        cell('A2')
                    }
                    cell('A3')
                    cell('A4')
                }
            }
        }

        when:
        CellRangeAddress range = workbook.getSheetAt(0).getMergedRegion(0)

        then:
        range.firstRow == 0
        range.lastRow == 1
        range.firstColumn == 1
        range.lastColumn == 1
    }

    void "test skipCells"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                column {
                    cell('A1')
                    cell('A2')
                    skipCells(2)
                    cell('A5')
                    cell('A6')
                }
            }
        }

        when:
        XSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.getRow(0).getCell(0).getStringCellValue() == 'A1'
        sheet.getRow(1).getCell(0).getStringCellValue() == 'A2'
        sheet.getRow(4).getCell(0).getStringCellValue() == 'A5'
        sheet.getRow(5).getCell(0).getStringCellValue() == 'A6'
    }
}
