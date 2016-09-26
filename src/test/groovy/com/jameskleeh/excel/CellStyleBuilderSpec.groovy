package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

import java.awt.Color

/**
 * Created by jameskleeh on 9/25/16.
 */
class CellStyleBuilderSpec extends Specification {

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
        options.font == [bold: true]

        when:
        options = [font: Font.UNDERLINE]
        cellStyleBuilder.convertSimpleOptions(options)

        then:
        options.font == [underline: (byte)1]
    }

    void "test buildStyle format"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle
        Excel.registerCellFormat(String, "bar")

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: "foo"])

        then:
        cellStyle.dataFormatString == "foo"

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: 1])

        then:
        cellStyle.dataFormat == (short)1

        when:
        cellStyle = cellStyleBuilder.buildStyle(null, [format: 1L])

        then:
        thrown(IllegalArgumentException)

        when:
        cellStyle = cellStyleBuilder.buildStyle("someString", [:])

        then:
        cellStyle.dataFormatString == "bar"
    }

    void "test buildStyle font"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: 1])

        then:
        thrown(IllegalArgumentException)

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
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: 'FFFFFF']])

        then:
        cellStyle.font.getXSSFColor().getRGB() == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: '#000000']])

        then:
        cellStyle.font.getXSSFColor().getRGB() == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: Color.BLUE]])

        then:
        cellStyle.font.getXSSFColor().getRGB() == [0, 0, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [font: [color: 1L]])

        then:
        thrown(IllegalArgumentException)
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
        cellStyle = cellStyleBuilder.buildStyle('', [hidden: "x"])

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
        cellStyle = cellStyleBuilder.buildStyle('', [locked: "x"])

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
        cellStyle = cellStyleBuilder.buildStyle('', [wrapped: "x"])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle vertical alignment"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: VerticalAlignment.BOTTOM])

        then:
        cellStyle.verticalAlignmentEnum == VerticalAlignment.BOTTOM

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: 'bottom'])

        then:
        cellStyle.verticalAlignmentEnum == VerticalAlignment.BOTTOM

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [verticalAlignment: 'x'])

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
        cellStyle = cellStyleBuilder.buildStyle('', [rotation: "foo"])

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
        cellStyle = cellStyleBuilder.buildStyle('', [indent: "foo"])

        then:
        thrown(ClassCastException)
    }

    void "test buildStyle border"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [border: [style: BorderStyle.DOTTED, color: Color.RED]])

        then:
        cellStyle.borderLeftEnum == BorderStyle.DOTTED
        cellStyle.leftBorderXSSFColor.getRGB() == [255, 0, 0] as byte[]
        cellStyle.borderRightEnum == BorderStyle.DOTTED
        cellStyle.rightBorderXSSFColor.getRGB() == [255, 0, 0] as byte[]
        cellStyle.borderBottomEnum == BorderStyle.DOTTED
        cellStyle.bottomBorderXSSFColor.getRGB() == [255, 0, 0] as byte[]
        cellStyle.borderTopEnum == BorderStyle.DOTTED
        cellStyle.topBorderXSSFColor.getRGB() == [255, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [border: [style: BorderStyle.DOTTED, color: Color.RED, left: [color: Color.BLUE], right: [style: BorderStyle.DASHED], bottom: [color: "FFFFFF"], top: [color: "#000000"]]])

        then:
        cellStyle.borderLeftEnum == BorderStyle.DOTTED
        cellStyle.leftBorderXSSFColor.getRGB() == [0, 0, 255] as byte[]
        cellStyle.borderRightEnum == BorderStyle.DASHED
        cellStyle.rightBorderXSSFColor.getRGB() == [255, 0, 0] as byte[]
        cellStyle.borderBottomEnum == BorderStyle.DOTTED
        cellStyle.bottomBorderXSSFColor.getRGB() == [255, 255, 255] as byte[]
        cellStyle.borderTopEnum == BorderStyle.DOTTED
        cellStyle.topBorderXSSFColor.getRGB() == [0, 0, 0] as byte[]
    }

    void "test buildStyle fill"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [:])

        then:
        cellStyle.fillPatternEnum == FillPatternType.SOLID_FOREGROUND

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: FillPatternType.DIAMONDS])

        then:
        cellStyle.fillPatternEnum == FillPatternType.DIAMONDS

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: "diamonds"])

        then:
        cellStyle.fillPatternEnum == FillPatternType.DIAMONDS

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [fill: "x"])

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
        cellStyle.fillForegroundXSSFColor.getRGB() == [255, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: "FFFFFF"])

        then:
        cellStyle.fillForegroundXSSFColor.getRGB() == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: "#000000"])

        then:
        cellStyle.fillForegroundXSSFColor.getRGB() == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: "blue"])

        then:
        thrown(IllegalArgumentException)
    }

    void "test buildStyle background color"() {
        given:
        CellStyleBuilder cellStyleBuilder = new CellStyleBuilder(new XSSFWorkbook())
        XSSFCellStyle cellStyle

        when: "Only the background color is specified"
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: Color.RED])

        then: "The foreground color is set instead of the background"
        cellStyle.fillForegroundXSSFColor.getRGB() == [255, 0, 0] as byte[]
        cellStyle.fillBackgroundXSSFColor == null

        when: "Both the foreground and background colors are specified"
        cellStyle = cellStyleBuilder.buildStyle('', [foregroundColor: Color.BLUE, backgroundColor: Color.RED])

        then: "Both are set"
        cellStyle.fillForegroundXSSFColor.getRGB() == [0, 0, 255] as byte[]
        cellStyle.fillBackgroundXSSFColor.getRGB() == [255, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: "FFFFFF"])

        then:
        cellStyle.fillForegroundXSSFColor.getRGB() == [255, 255, 255] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: "#000000"])

        then:
        cellStyle.fillForegroundXSSFColor.getRGB() == [0, 0, 0] as byte[]

        when:
        cellStyle = cellStyleBuilder.buildStyle('', [backgroundColor: "blue"])

        then:
        thrown(IllegalArgumentException)
    }

}
