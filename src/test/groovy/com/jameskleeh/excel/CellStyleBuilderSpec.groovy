package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

import java.awt.Color

/**
 * Created by jameskleeh on 9/25/16.
 */
class CellStyleBuilderSpec extends Specification {

    void cleanup() {
        Excel.formatEntries.clear()
        Excel.rendererEntries.clear()
    }

    void "test convertSimpleOptions"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        Map<String, Object> options = [border: BorderStyle.DASHED]

        when:
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.border == [style: BorderStyle.DASHED]

        when:
        options = [border: [style: BorderStyle.DASHED, left: BorderStyle.DOTTED]]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.border == [style: BorderStyle.DASHED, left: [style: BorderStyle.DOTTED]]

        when:
        options = [font: Font.BOLD]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == [bold: true, italic: false, strikeout: false, underline: (byte)0]

        when:
        options = [font: Font.ITALIC]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == [bold: false, italic: true, strikeout: false, underline: (byte)0]

        when:
        options = [font: Font.STRIKEOUT]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == [bold: false, italic: false, strikeout: true, underline: (byte)0]

        when:
        options = [font: Font.UNDERLINE]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == [bold: false, italic: false, strikeout: false, underline: (byte)1]

        when:
        options = [font: 'x']
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == 'x'
    }

    void "test buildStyle format"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle
        Excel.registerCellFormat(String, 'bar')

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: 'foo'])

        then:
        cellStyle.dataFormatString == 'foo'

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: 1])

        then:
        cellStyle.dataFormat == (short)1

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: 1L])

        then:
        thrown(IllegalArgumentException)

        when: 'A call to getStyle is essential here to bring in the formats registered in Excel'
        cellStyle = cellStyleBuilder.getStyle('someString', [:])

        then:
        cellStyle.dataFormatString == 'bar'
    }

    void "test buildStyle font"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: 1])

        then:
        Exception ex = thrown(IllegalArgumentException)
        ex.message == 'The font option must be an instance of a Map'

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [:]])

        then:
        !cellStyle.font.bold
        !cellStyle.font.italic
        !cellStyle.font.strikeout
        cellStyle.font.underline == (byte)0

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [bold: true, italic: true, strikeout: true, underline: true]])

        then:
        cellStyle.font.bold
        cellStyle.font.italic
        cellStyle.font.strikeout
        cellStyle.font.underline == (byte)1

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [bold: 'foo']])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 'single']])

        then:
        cellStyle.font.underline == (byte)1

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 'singleAccounting']])

        then:
        cellStyle.font.underline == (byte)0x21

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 'double']])

        then:
        cellStyle.font.underline == (byte)2

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 'doubleAccounting']])

        then:
        cellStyle.font.underline == (byte)0x22

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 'foo']])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [underline: 1]])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: 'FFFFFF']])

        then:
        cellStyle.font.XSSFColor.RGB == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: '#000000']])

        then:
        cellStyle.font.XSSFColor.RGB == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: Color.BLUE]])

        then:
        cellStyle.font.XSSFColor.RGB == [0, 0, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: 1L]])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [size: 13]])

        then:
        cellStyle.font.fontHeight == (short)260

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [size: 12.5]])

        then:
        cellStyle.font.fontHeight == (short)250

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [name: 'Arial']])

        then:
        cellStyle.font.fontName == 'Arial'
    }

    void "test buildStyle hidden"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [hidden: true])

        then:
        cellStyle.hidden

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [hidden: false])

        then:
        !cellStyle.hidden

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [hidden: 'x'])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle locked"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [locked: true])

        then:
        cellStyle.locked

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [locked: false])

        then:
        !cellStyle.locked

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [locked: 'x'])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle wrapped"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [wrapped: true])

        then:
        cellStyle.wrapText

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [wrapped: false])

        then:
        !cellStyle.wrapText

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [wrapped: 'x'])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle horizontal alignment"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [alignment: HorizontalAlignment.RIGHT])

        then:
        cellStyle.alignmentEnum == HorizontalAlignment.RIGHT

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [alignment: 'right'])

        then:
        cellStyle.alignmentEnum == HorizontalAlignment.RIGHT

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [alignment: 'x'])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [alignment: 1])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle vertical alignment"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: VerticalAlignment.TOP])

        then:
        cellStyle.verticalAlignmentEnum == VerticalAlignment.TOP

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: 'top'])

        then:
        cellStyle.verticalAlignmentEnum == VerticalAlignment.TOP

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: 'x'])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: 1])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle rotation"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [rotation: 0])

        then:
        cellStyle.rotation == (short)0

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [rotation: 1L])

        then:
        cellStyle.rotation == (short)1

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [rotation: 'foo'])

        then:
        thrown(ClassCastException)
    }

    void "test buildStyle indention"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [indent: 0])

        then:
        cellStyle.indention == (short)0

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [indent: 1L])

        then:
        cellStyle.indention == (short)1

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [indent: 'foo'])

        then:
        thrown(ClassCastException)
    }

    void "test buildStyle border"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyleBuilder.buildStyle('', [border: [style: 1]])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [border: [style: BorderStyle.DOTTED, color: Color.RED]])
        byte[] color = [255, 0, 0] as byte[]

        then:
        cellStyle.borderLeftEnum == BorderStyle.DOTTED
        cellStyle.leftBorderXSSFColor.RGB == color
        cellStyle.borderRightEnum == BorderStyle.DOTTED
        cellStyle.rightBorderXSSFColor.RGB == color
        cellStyle.borderBottomEnum == BorderStyle.DOTTED
        cellStyle.bottomBorderXSSFColor.RGB == color
        cellStyle.borderTopEnum == BorderStyle.DOTTED
        cellStyle.topBorderXSSFColor.RGB == color

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [border: [style: BorderStyle.DOTTED, color: Color.RED, left: [color: Color.BLUE], right: [style: BorderStyle.DASHED], bottom: [color: 'FFFFFF'], top: [color: '#000000']]])

        then:
        cellStyle.borderLeftEnum == BorderStyle.DOTTED
        cellStyle.leftBorderXSSFColor.RGB == [0, 0, 255] as byte[]
        cellStyle.borderRightEnum == BorderStyle.DASHED
        cellStyle.rightBorderXSSFColor.RGB == [255, 0, 0] as byte[]
        cellStyle.borderBottomEnum == BorderStyle.DOTTED
        cellStyle.bottomBorderXSSFColor.RGB == [255, 255, 255] as byte[]
        cellStyle.borderTopEnum == BorderStyle.DOTTED
        cellStyle.topBorderXSSFColor.RGB == [0, 0, 0] as byte[]

        when:
        Map options = [border: [style: BorderStyle.DOTTED, left: BorderStyle.THIN]]
        cellStyleBuilder.convertSimpleOptions(options)
        cellStyle = cellStyleBuilder.buildStyle('', options)

        then:
        cellStyle.borderLeftEnum == BorderStyle.THIN
        cellStyle.borderRightEnum == BorderStyle.DOTTED
        cellStyle.borderBottomEnum == BorderStyle.DOTTED
        cellStyle.borderTopEnum == BorderStyle.DOTTED
    }

    void "test buildStyle fill"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [:])

        then:
        cellStyle.fillPatternEnum == FillPatternType.NO_FILL

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: FillPatternType.DIAMONDS])

        then:
        cellStyle.fillPatternEnum == FillPatternType.DIAMONDS

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: 'diamonds'])

        then:
        cellStyle.fillPatternEnum == FillPatternType.DIAMONDS

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: 'x'])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: 1])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle foreground color"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: Color.RED])

        then:
        cellStyle.fillForegroundXSSFColor.RGB == [255, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: 'FFFFFF'])

        then:
        cellStyle.fillForegroundXSSFColor.RGB == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: '#000000'])

        then:
        cellStyle.fillForegroundXSSFColor.RGB == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: 'blue'])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle background color"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when: 'Only the background color is specified'
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: Color.RED])

        then: 'The foreground color is set instead of the background and the fill pattern is set to solid'
        cellStyle.fillForegroundXSSFColor.RGB == [255, 0, 0] as byte[]
        cellStyle.fillBackgroundXSSFColor == null
        cellStyle.fillPatternEnum == FillPatternType.SOLID_FOREGROUND

        when: 'Both the foreground and background colors are specified'
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: Color.BLUE, backgroundColor: Color.RED])

        then: 'Both are set'
        cellStyle.fillForegroundXSSFColor.RGB == [0, 0, 255] as byte[]
        cellStyle.fillBackgroundXSSFColor.RGB == [255, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: 'FFFFFF'])

        then:
        cellStyle.fillForegroundXSSFColor.RGB == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: '#000000'])

        then:
        cellStyle.fillForegroundXSSFColor.RGB == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: 'blue'])

        then:
        thrown(IllegalArgumentException)
    }

    void "test getStyle pulls from cache"() {
        XSSFWorkbook workbook = new XSSFWorkbook()
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(workbook)

        when:
        cellStyleBuilder.getStyle('', [hidden: true])
        cellStyleBuilder.getStyle('', [hidden: true])

        then:
        cellStyleBuilder.workbookCache.containsStyle([hidden: true])
        cellStyleBuilder.workbookCache.styles.size() == 1
    }

    void "test setStyle cell no options"() {
        given:
        XSSFWorkbook workbook = new XSSFWorkbook()
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(workbook)
        XSSFCell defaultCell = workbook.createSheet().createRow(0).createCell(0)

        when:
        XSSFCell cell = defaultCell.row.createCell(1)
        cellStyleBuilder.setStyle('', cell, null)

        then:
        cell.cellStyle == defaultCell.cellStyle
    }

    void "test setStyle cell options"() {
        given:
        XSSFWorkbook workbook = new XSSFWorkbook()
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(workbook)
        XSSFCell defaultCell = workbook.createSheet().createRow(0).createCell(0)

        when:
        XSSFCell cell = defaultCell.row.createCell(1)
        cellStyleBuilder.setStyle('', cell, [font: Font.ITALIC])

        then:
        cell.cellStyle.font.italic
    }

    void "test setStyle cell options are merged"() {
        given:
        XSSFWorkbook workbook = new XSSFWorkbook()
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(workbook)
        XSSFCell defaultCell = workbook.createSheet().createRow(0).createCell(0)

        when:
        XSSFCell cell = defaultCell.row.createCell(1)
        cellStyleBuilder.setStyle('', cell, [font: Font.ITALIC, border: [left: BorderStyle.DOUBLE]], [font: Font.BOLD, border: BorderStyle.DASH_DOT])
        XSSFCellStyle style = cell.cellStyle

        then:
        style.font.italic
        !style.font.bold
        style.borderLeftEnum == BorderStyle.DOUBLE
        style.borderTopEnum == BorderStyle.DASH_DOT
        style.borderRightEnum == BorderStyle.DASH_DOT
        style.borderBottomEnum == BorderStyle.DASH_DOT
    }

    void "test merging of options with rows"() {
        given:
        XSSFCell testCellLeft
        XSSFCell testCellMiddle
        XSSFCell testCellMiddle2
        XSSFCell testCellRight
        XSSFCell testCell2
        XSSFCell testCell3
        XSSFCell testCell4

        ExcelBuilder.build {
            sheet {
                defaultStyle([border: BorderStyle.MEDIUM])
                row {
                    defaultStyle([border: [left: BorderStyle.HAIR]])
                    merge([border: [right: BorderStyle.THICK], alignment: 'center']) {
                        testCellLeft = cell('Foo')
                        testCellMiddle = cell()
                        testCellMiddle2 = cell()
                        testCellRight = cell()
                    }
                    testCell2 = cell('')
                }
                row {
                    testCell3 = cell('')
                }
            }
            sheet {
                row {
                    testCell4 = cell('')
                }
            }
        }

        when:
        XSSFCellStyle mergeLeft = testCellLeft.cellStyle //sheet, row, merge
        XSSFCellStyle mergeMiddle = testCellMiddle.cellStyle //sheet, row, merge
        XSSFCellStyle mergeMiddle2 = testCellMiddle2.cellStyle //sheet, row, merge
        XSSFCellStyle mergeRight = testCellRight.cellStyle //sheet, row, merge
        XSSFCellStyle style2 = testCell2.cellStyle //sheet, row
        XSSFCellStyle style3 = testCell3.cellStyle //sheet
        XSSFCellStyle style4 = testCell4.cellStyle //none

        then:
        mergeLeft.borderLeftEnum == BorderStyle.HAIR
        mergeLeft.borderTopEnum == BorderStyle.MEDIUM
        mergeLeft.borderRightEnum == BorderStyle.NONE
        mergeLeft.borderBottomEnum == BorderStyle.MEDIUM

        mergeMiddle.borderLeftEnum == BorderStyle.NONE
        mergeMiddle.borderRightEnum == BorderStyle.NONE
        mergeMiddle.borderTopEnum == BorderStyle.MEDIUM
        mergeMiddle.borderBottomEnum == BorderStyle.MEDIUM

        mergeMiddle2.borderLeftEnum == BorderStyle.NONE
        mergeMiddle2.borderRightEnum == BorderStyle.NONE
        mergeMiddle2.borderTopEnum == BorderStyle.MEDIUM
        mergeMiddle2.borderBottomEnum == BorderStyle.MEDIUM

        mergeRight.borderLeftEnum == BorderStyle.NONE
        mergeRight.borderTopEnum == BorderStyle.MEDIUM
        mergeRight.borderRightEnum == BorderStyle.THICK
        mergeRight.borderBottomEnum == BorderStyle.MEDIUM

        style2.borderLeftEnum == BorderStyle.HAIR
        style2.borderTopEnum == BorderStyle.MEDIUM
        style2.borderRightEnum == BorderStyle.MEDIUM
        style2.borderBottomEnum == BorderStyle.MEDIUM
        style3.borderLeftEnum == BorderStyle.MEDIUM
        style3.borderTopEnum == BorderStyle.MEDIUM
        style3.borderRightEnum == BorderStyle.MEDIUM
        style3.borderBottomEnum == BorderStyle.MEDIUM
        style4.borderLeftEnum == BorderStyle.NONE
        style4.borderTopEnum == BorderStyle.NONE
        style4.borderRightEnum == BorderStyle.NONE
        style4.borderBottomEnum == BorderStyle.NONE
    }

    void "test merging of options with columns"() {
        given:
        XSSFCell testCellTop
        XSSFCell testCellMiddle
        XSSFCell testCellMiddle2
        XSSFCell testCellBottom
        XSSFCell testCell2
        XSSFCell testCell3
        XSSFCell testCell4
        ExcelBuilder.build {
            sheet {
                defaultStyle([border: BorderStyle.MEDIUM])
                column {
                    defaultStyle([border: [left: BorderStyle.HAIR]])
                    merge([border: [right: BorderStyle.THICK], alignment: 'center']) {
                        testCellTop = cell('Foo')
                        testCellMiddle = cell()
                        testCellMiddle2 = cell()
                        testCellBottom = cell()
                    }
                    testCell2 = cell('')
                }
                column {
                    testCell3 = cell('')
                }
            }
            sheet {
                column {
                    testCell4 = cell('')
                }
            }
        }

        when:
        XSSFCellStyle mergeTop = testCellTop.cellStyle //sheet, row, merge
        XSSFCellStyle mergeMiddle = testCellMiddle.cellStyle //sheet, row, merge
        XSSFCellStyle mergeMiddle2 = testCellMiddle2.cellStyle //sheet, row, merge
        XSSFCellStyle mergeBottom = testCellBottom.cellStyle //sheet, row, merge
        XSSFCellStyle style2 = testCell2.cellStyle //sheet, row
        XSSFCellStyle style3 = testCell3.cellStyle //sheet
        XSSFCellStyle style4 = testCell4.cellStyle //none

        then:
        mergeTop.borderLeftEnum == BorderStyle.HAIR
        mergeTop.borderTopEnum == BorderStyle.MEDIUM
        mergeTop.borderRightEnum == BorderStyle.THICK
        mergeTop.borderBottomEnum == BorderStyle.NONE

        mergeMiddle.borderLeftEnum == BorderStyle.HAIR
        mergeMiddle.borderRightEnum == BorderStyle.THICK
        mergeMiddle.borderTopEnum == BorderStyle.NONE
        mergeMiddle.borderBottomEnum == BorderStyle.NONE

        mergeMiddle2.borderLeftEnum == BorderStyle.HAIR
        mergeMiddle2.borderRightEnum == BorderStyle.THICK
        mergeMiddle2.borderTopEnum == BorderStyle.NONE
        mergeMiddle2.borderBottomEnum == BorderStyle.NONE

        mergeBottom.borderLeftEnum == BorderStyle.HAIR
        mergeBottom.borderTopEnum == BorderStyle.NONE
        mergeBottom.borderRightEnum == BorderStyle.THICK
        mergeBottom.borderBottomEnum == BorderStyle.MEDIUM

        style2.borderLeftEnum == BorderStyle.HAIR
        style2.borderTopEnum == BorderStyle.MEDIUM
        style2.borderRightEnum == BorderStyle.MEDIUM
        style2.borderBottomEnum == BorderStyle.MEDIUM
        style3.borderLeftEnum == BorderStyle.MEDIUM
        style3.borderTopEnum == BorderStyle.MEDIUM
        style3.borderRightEnum == BorderStyle.MEDIUM
        style3.borderBottomEnum == BorderStyle.MEDIUM
        style4.borderLeftEnum == BorderStyle.NONE
        style4.borderTopEnum == BorderStyle.NONE
        style4.borderRightEnum == BorderStyle.NONE
        style4.borderBottomEnum == BorderStyle.NONE
    }

    void "test number of styles created"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row {
                    merge([font: [color: Color.YELLOW], border: [style: BorderStyle.DOTTED, color: Color.RED, left: [color: Color.BLUE], right: [style: BorderStyle.DASHED], bottom: [color: '7900bf'], top: [color: '#2AB54A']]]) {
                        cell('Test')
                        skipCells(3)
                    }
                    cell('another', [font: [color: Color.YELLOW]])
                }
            }
            sheet {
                column {
                    merge([font: [color: Color.YELLOW], border: [style: BorderStyle.DOTTED, color: Color.RED, left: [color: Color.BLUE], right: [style: BorderStyle.DASHED], bottom: [color: '7900bf'], top: [color: '#2AB54A']]]) {
                        cell('Test')
                        skipCells(3)
                    }
                    cell('another', [font: [color: Color.YELLOW]])
                }
            }
        }

        expect:
        //3 for each merge = 6 + default + yellow color
        workbook.numCellStyles == 8
    }

}
