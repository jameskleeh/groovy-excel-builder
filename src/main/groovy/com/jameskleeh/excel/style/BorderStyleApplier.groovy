package com.jameskleeh.excel.style

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide

/**
 * Created by jameskleeh on 7/3/17.
 */
interface BorderStyleApplier {

    void applyStyle(BorderSide side, BorderStyle style)

    void applyStyle(BorderStyle style)

    void applyColor(BorderSide side, XSSFColor color)

    void applyColor(XSSFColor color)
}