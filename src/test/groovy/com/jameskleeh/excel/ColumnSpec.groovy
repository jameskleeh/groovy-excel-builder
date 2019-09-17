package com.jameskleeh.excel

import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Specification

class ColumnSpec extends Specification {

    void "test output by column"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                columns {
                    column('Column A1', 'id')
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
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.getRow(0).getCell(0).stringCellValue == 'Column A1'
        sheet.getRow(0).getCell(1) == null
        sheet.getRow(1).getCell(0).stringCellValue == 'A2'
        sheet.getRow(1).getCell(1).stringCellValue == 'B2'
        sheet.getRow(2).getCell(0).stringCellValue == 'A3'
        sheet.getRow(2).getCell(1).stringCellValue == 'B3'

    }

    void "test merge"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
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
        SXSSFWorkbook workbook = ExcelBuilder.build {
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
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.getRow(0).getCell(0).stringCellValue == 'A1'
        sheet.getRow(1).getCell(0).stringCellValue == 'A2'
        sheet.getRow(4).getCell(0).stringCellValue == 'A5'
        sheet.getRow(5).getCell(0).stringCellValue == 'A6'
    }

    void "test link"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet('X') {
                row {
                    link('Test URL', 'http://www.google.com', HyperlinkType.URL)
                    link('Test File', 'test.docx', HyperlinkType.FILE)
                    link('Test Email', 'mailto:foo@bar.com', HyperlinkType.EMAIL)
                    link('Test Document') {
                        "'${sheetName}'!${exactCell(1, 1)}"
                    }
                }
            }
        }

        when:
        SXSSFRow row = workbook.getSheetAt(0).getRow(0)
        List<Cell> cells = row.cellIterator().toList()

        then:
        cells[0].stringCellValue == 'Test URL'
        cells[0].hyperlink.address == 'http://www.google.com'
        cells[0].hyperlink.type == HyperlinkType.URL
        cells[1].stringCellValue == 'Test File'
        cells[1].hyperlink.address == 'test.docx'
        cells[1].hyperlink.type == HyperlinkType.FILE
        cells[2].stringCellValue == 'Test Email'
        cells[2].hyperlink.address == 'mailto:foo@bar.com'
        cells[2].hyperlink.type == HyperlinkType.EMAIL
        cells[3].stringCellValue == 'Test Document'
        cells[3].hyperlink.address == "'X'!B2"
        cells[3].hyperlink.type == HyperlinkType.DOCUMENT
    }
}
