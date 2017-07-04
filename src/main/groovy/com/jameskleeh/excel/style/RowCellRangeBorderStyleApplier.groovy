package com.jameskleeh.excel.style

import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.BOTTOM
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.LEFT
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.RIGHT
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.TOP
import groovy.transform.InheritConstructors
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

/**
 * Applys styles and borders to a single row of merged cells.
 *
 * @author James Kleeh
 */
@InheritConstructors
class RowCellRangeBorderStyleApplier extends CellRangeBorderStyleApplier {

    @Override
    void applyStyle(XSSFCellBorder.BorderSide side, BorderStyle style) {
        switch (side) {
            case TOP:
                leftTop.setBorderTop(style)
                bottomRight.setBorderTop(style)
                middle?.setBorderTop(style)
                break
            case BOTTOM:
                leftTop.setBorderBottom(style)
                bottomRight.setBorderBottom(style)
                middle?.setBorderBottom(style)
                break
            case LEFT:
                leftTop.setBorderLeft(style)
                break
            case RIGHT:
                bottomRight.setBorderRight(style)
                break
        }
    }

    @Override
    void applyColor(XSSFCellBorder.BorderSide side, XSSFColor color) {
        switch (side) {
            case TOP:
                leftTop.setTopBorderColor(color)
                bottomRight.setTopBorderColor(color)
                middle?.setTopBorderColor(color)
                break
            case BOTTOM:
                leftTop.setBottomBorderColor(color)
                bottomRight.setBottomBorderColor(color)
                middle?.setBottomBorderColor(color)
                break
            case LEFT:
                leftTop.setLeftBorderColor(color)
                break
            case RIGHT:
                bottomRight.setRightBorderColor(color)
                break
        }
    }
}
