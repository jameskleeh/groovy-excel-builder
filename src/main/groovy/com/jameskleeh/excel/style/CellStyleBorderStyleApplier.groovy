package com.jameskleeh.excel.style

import groovy.transform.CompileStatic
import groovy.transform.TupleConstructor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.*

@CompileStatic
@TupleConstructor
class CellStyleBorderStyleApplier implements BorderStyleApplier {

    XSSFCellStyle cellStyle

    @Override
    void applyStyle(BorderSide side, BorderStyle style) {
        switch(side) {
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
        switch(side) {
            case TOP:
                cellStyle.setTopBorderColor(color)
                break
            case BOTTOM:
                cellStyle.setBottomBorderColor(color)
                break
            case LEFT:
                cellStyle.setLeftBorderColor(color)
                break
            case RIGHT:
                cellStyle.setRightBorderColor(color)
                break
        }
    }

    @Override
    void applyColor(XSSFColor color) {
        cellStyle.setTopBorderColor(color)
        cellStyle.setBottomBorderColor(color)
        cellStyle.setLeftBorderColor(color)
        cellStyle.setRightBorderColor(color)
    }

}
