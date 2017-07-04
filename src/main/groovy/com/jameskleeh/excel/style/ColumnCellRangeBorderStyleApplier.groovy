package com.jameskleeh.excel.style

import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.BOTTOM
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.LEFT
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.RIGHT
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.TOP
import groovy.transform.CompileStatic
import groovy.transform.InheritConstructors
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

/**
 * Applys styles and borders to a single column of merged cells.
 *
 * @author James Kleeh
 */
@InheritConstructors
@CompileStatic
class ColumnCellRangeBorderStyleApplier extends CellRangeBorderStyleApplier {

    @Override
    void applyStyle(XSSFCellBorder.BorderSide side, BorderStyle style) {
        switch (side) {
            case TOP:
                leftTop.setBorderTop(style)
                break
            case BOTTOM:
                bottomRight.setBorderBottom(style)
                break
            case LEFT:
                leftTop.setBorderLeft(style)
                bottomRight.setBorderLeft(style)
                middle?.setBorderLeft(style)
                break
            case RIGHT:
                leftTop.setBorderRight(style)
                bottomRight.setBorderRight(style)
                middle?.setBorderRight(style)
                break
        }
    }

    @Override
    void applyColor(XSSFCellBorder.BorderSide side, XSSFColor color) {
        switch (side) {
            case TOP:
                leftTop.setTopBorderColor(color)
                break
            case BOTTOM:
                bottomRight.setBottomBorderColor(color)
                break
            case LEFT:
                leftTop.setLeftBorderColor(color)
                bottomRight.setLeftBorderColor(color)
                middle?.setLeftBorderColor(color)
                break
            case RIGHT:
                leftTop.setRightBorderColor(color)
                bottomRight.setRightBorderColor(color)
                middle?.setRightBorderColor(color)
                break
        }
    }
}
