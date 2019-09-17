package com.jameskleeh.excel.style

import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.*

import groovy.transform.CompileStatic
import groovy.transform.TupleConstructor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide

/**
 * Applys styles and borders to a cell style object
 *
 * @author James Kleeh
 */
@CompileStatic
@TupleConstructor
@SuppressWarnings('NoWildcardImports')
class CellStyleBorderStyleApplier implements BorderStyleApplier {

    XSSFCellStyle cellStyle

    @Override
    void applyStyle(BorderSide side, BorderStyle style) {
        switch (side) {
            case TOP:
                cellStyle.setBorderTop(style)
                break
            case BOTTOM:
                cellStyle.setBorderBottom(style)
                break
            case LEFT:
                cellStyle.setBorderLeft(style)
                break
            case RIGHT:
                cellStyle.setBorderRight(style)
                break
        }
    }

    @Override
    void applyStyle(BorderStyle style) {
        cellStyle.setBorderTop(style)
        cellStyle.setBorderBottom(style)
        cellStyle.setBorderLeft(style)
        cellStyle.setBorderRight(style)
    }

    @Override
    void applyColor(BorderSide side, XSSFColor color) {
        cellStyle.setBorderColor(side, color)
    }

    @Override
    void applyColor(XSSFColor color) {
        cellStyle.setTopBorderColor(color)
        cellStyle.setBottomBorderColor(color)
        cellStyle.setLeftBorderColor(color)
        cellStyle.setRightBorderColor(color)
    }

}
