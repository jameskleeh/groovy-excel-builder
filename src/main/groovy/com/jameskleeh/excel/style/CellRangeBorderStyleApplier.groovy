package com.jameskleeh.excel.style

import groovy.transform.CompileStatic
import groovy.transform.TupleConstructor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide
import static org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide.*

@CompileStatic
@TupleConstructor
class CellRangeBorderStyleApplier implements BorderStyleApplier {

    CellRangeAddress range
    Sheet sheet

    @Override
    void applyStyle(BorderSide side, BorderStyle style) {
        switch(side) {
            case TOP:
                RegionUtil.setBorderTop(style, range, sheet)
                break
            case BOTTOM:
                RegionUtil.setBorderBottom(style, range, sheet)
                break
            case LEFT:
                RegionUtil.setBorderLeft(style, range, sheet)
                break
            case RIGHT:
                RegionUtil.setBorderRight(style, range, sheet)
                break
        }
    }

    @Override
    void applyStyle(BorderStyle style) {
        RegionUtil.setBorderTop(style, range, sheet)
        RegionUtil.setBorderBottom(style, range, sheet)
        RegionUtil.setBorderLeft(style, range, sheet)
        RegionUtil.setBorderRight(style, range, sheet)
    }

    @Override
    void applyColor(BorderSide side, XSSFColor color) {
        switch(side) {
            case TOP:
                RegionUtil.setTopBorderColor(color.index, range, sheet)
                break
            case BOTTOM:
                RegionUtil.setBottomBorderColor(color.index, range, sheet)
                break
            case LEFT:
                RegionUtil.setLeftBorderColor(color.index, range, sheet)
                break
            case RIGHT:
                RegionUtil.setRightBorderColor(color.index, range, sheet)
                break
        }
    }

    @Override
    void applyColor(XSSFColor color) {
        RegionUtil.setTopBorderColor(color.index, range, sheet)
        RegionUtil.setBottomBorderColor(color.index, range, sheet)
        RegionUtil.setLeftBorderColor(color.index, range, sheet)
        RegionUtil.setRightBorderColor(color.index, range, sheet)
    }
}
