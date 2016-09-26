package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

/**
 * Created by jameskleeh on 9/25/16.
 */
class RowSpec extends Specification {

    void "test skipCells"() {
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                row {
                    cell()
                    skipCells(2)
                    cell()
                }
            }
        }

        when:
        Iterator<Cell> cells = workbook.getSheetAt(0).getRow(0).cellIterator()

        then:
        cells.next().columnIndex == 0
        cells.next().columnIndex == 3
        !cells.hasNext()
    }

    void "test skipTo"() {
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                columns {
                    column("Foo", "foo")
                    skipCells(2)
                    column("Bar", "bar")
                }
                row {
                    skipTo("bar")
                    cell()
                }
            }
        }

        when:
        Iterator<Cell> cells = workbook.getSheetAt(0).getRow(1).cellIterator()

        then:
        cells.next().columnIndex == 3
        !cells.hasNext()
    }

    void "test formula(String)"() {
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                row {
                    formula("=SUM()")
                    formula("SUM()")
                    formula {
                        "=CONCATENATE()"
                    }
                    formula {
                        "CONCATENATE()"
                    }
                }
            }
        }

        when:
        Iterator<Cell> cells = workbook.getSheetAt(0).getRow(0).cellIterator()

        then:
        cells.next().cellFormula == "SUM()"
        cells.next().cellFormula == "SUM()"
        cells.next().cellFormula == "CONCATENATE()"
        cells.next().cellFormula == "CONCATENATE()"
        !cells.hasNext()
    }

    void "test cell"() {
        Excel.registerCellRenderer(StringBuilder) {
            it.append('x').toString()
        }
        XSSFWorkbook workbook = new ExcelBuilder().build {
            sheet {
                row {
                    cell()
                    cell("A")
                    cell(Calendar.instance)
                    cell(new Date())
                    cell(new Double(2.2))
                    cell(false)
                    cell(new StringBuilder('foo'))
                    cell(1L)
                }
            }
        }

        when:
        Iterator<Cell> cells = workbook.getSheetAt(0).getRow(0).cellIterator()

        then:
        cells.next().stringCellValue == ''
        cells.next().stringCellValue == 'A'
        cells.next().dateCellValue.clearTime() == new Date().clearTime()
        cells.next().dateCellValue.clearTime() == new Date().clearTime()
        cells.next().numericCellValue == new Double(2.2)
        cells.next().booleanCellValue == false
        cells.next().stringCellValue == 'foox'
        cells.next().stringCellValue == "1"
    }
}
