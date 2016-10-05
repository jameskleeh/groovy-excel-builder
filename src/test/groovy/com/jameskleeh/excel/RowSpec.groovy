package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

class RowSpec extends Specification {

    void "test skipCells"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
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
        XSSFWorkbook workbook = ExcelBuilder.build {
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

    void "test skipTo overwrite previous cells"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                columns {
                    column("Foo", "foo")
                    skipCells(2)
                    column("Bar", "bar")
                }
                row {
                    cell()
                    cell()
                    skipTo("foo")
                    cell("A1")
                    cell("A2")
                }
            }
        }

        when:
        Iterator<Cell> cells = workbook.getSheetAt(0).getRow(1).cellIterator()

        then:
        cells.next().stringCellValue == 'A1'
        cells.next().stringCellValue == 'A2'
        !cells.hasNext()
    }

    void "test formula(String)"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
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
        XSSFWorkbook workbook = ExcelBuilder.build {
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
        cells.next().numericCellValue == new Double(1)
    }
}
