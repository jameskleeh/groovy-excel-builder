/*
Licensed to the Apache Software Foundation (ASF) under one
or more contributor license agreements.  See the NOTICE file
distributed with this work for additional information
regarding copyright ownership.  The ASF licenses this file
to you under the Apache License, Version 2.0 (the
"License"); you may not use this file except in compliance
with the License.  You may obtain a copy of the License at

  http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing,
software distributed under the License is distributed on an
"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, either express or implied.  See the License for the
specific language governing permissions and limitations
under the License.
*/
package com.jameskleeh.excel

import com.jameskleeh.excel.style.BorderStyleApplier
import com.jameskleeh.excel.style.CellRangeBorderStyleApplier
import com.jameskleeh.excel.style.CellStyleBorderStyleApplier
import groovy.transform.CompileStatic
import groovy.transform.TypeCheckingMode
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font as FontType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide

import java.awt.*

/**
 * A class to build an {@link org.apache.poi.xssf.usermodel.XSSFCellStyle} from a map
 *
 * @author James Kleeh
 * @since 0.1.0
 */
@CompileStatic
class CellStyleBuilder {

    private final SXSSFWorkbook workbook
    private final WorkbookCache workbookCache

    protected static final String FORMAT = 'format'
    protected static final String HIDDEN = 'hidden'
    protected static final String LOCKED = 'locked'
    protected static final String HORIZONTAL_ALIGNMENT = 'alignment'
    protected static final String VERTICAL_ALIGNMENT = 'verticalAlignment'
    protected static final String WRAPPED = 'wrapped'
    protected static final String ROTATION = 'rotation'
    protected static final String INDENT = 'indent'
    protected static final String BORDER = 'border'
    protected static final String LEFT = 'left'
    protected static final String RIGHT = 'right'
    protected static final String TOP = 'top'
    protected static final String BOTTOM = 'bottom'
    protected static final String COLOR = 'color'
    protected static final String STYLE = 'style'
    protected static final String FONT = 'font'
    protected static final String FONT_BOLD = 'bold'
    protected static final String FONT_ITALIC = 'italic'
    protected static final String FONT_STRIKEOUT = 'strikeout'
    protected static final String FONT_UNDERLINE = 'underline'
    protected static final String FONT_COLOR = 'color'
    protected static final String FONT_NAME = 'name'
    protected static final String FONT_SIZE = 'size'
    protected static final String FILL = 'fill'
    protected static final String BACKGROUND_COLOR = 'backgroundColor'
    protected static final String FOREGROUND_COLOR = 'foregroundColor'

    CellStyleBuilder(SXSSFWorkbook workbook) {
        this.workbook = workbook
        workbookCache = new WorkbookCache(workbook)
    }

    private static void convertBorderOptions(Map options, String key) {
        if (options.containsKey(key) && options[key] instanceof BorderStyle) {
            BorderStyle border = (BorderStyle)options.remove(key)
            options.put(key, [style: border])
        }
    }

    /**
     *
     * A method to convert global options into specific options.
     * Example:
     * [border: BorderStyle.THIN] would convert to
     * [border: [style: BorderStyle.THIN, left: [style: BorderStyle.THIN], right: ...]]
     *
     * @param options A map of options
     */
     static void convertSimpleOptions(Map options) {
        if (options) {
            if (options.containsKey(BORDER) && options[BORDER] instanceof BorderStyle) {
                BorderStyle border = (BorderStyle)options.remove(BORDER)
                options.put(BORDER, [style: border])
            }
            if (options.containsKey(FONT) && options[FONT] instanceof Font) {
                Font font = (Font)options.remove(FONT)
                Map fontOptions = [:]
                fontOptions[FONT_BOLD] = (font == Font.BOLD)
                fontOptions[FONT_ITALIC] = (font == Font.ITALIC)
                fontOptions[FONT_STRIKEOUT] = (font == Font.STRIKEOUT)
                fontOptions[FONT_UNDERLINE] = (font == Font.UNDERLINE ? (byte)1 : (byte)0)
                options[FONT] = fontOptions
            }
            if (options.containsKey(BORDER)) {
                Map border = (Map)options[BORDER]
                convertBorderOptions(border, LEFT)
                convertBorderOptions(border, RIGHT)
                convertBorderOptions(border, TOP)
                convertBorderOptions(border, BOTTOM)
            }
        }
    }

    private void setFormat(XSSFCellStyle cellStyle, Object format) {
        if (format instanceof Integer) {
            cellStyle.setDataFormat((Integer)format)
        } else if (format instanceof String) {
            cellStyle.setDataFormat(workbook.creationHelper.createDataFormat().getFormat((String)format))
        } else {
            throw new IllegalArgumentException('The cell format must be an Integer or String')
        }
    }

    private void setBooleanFont(Map options, String key, Closure callable) {
        if (options.containsKey(key)) {
            if (options[key] instanceof Boolean) {
                callable.call((Boolean)options[key])
            } else {
                throw new IllegalArgumentException("[font: [$key: <>]] must be a boolean")
            }
        }
    }

    private void setFont(XSSFCellStyle cellStyle, Object fontOptions) {
        if (!workbookCache.containsFont(fontOptions)) {
            XSSFFont font = (XSSFFont) workbook.createFont()
            if (fontOptions instanceof Map) {
                Map fontMap = (Map)fontOptions
                setBooleanFont(fontMap, FONT_BOLD, font.&setBold)
                setBooleanFont(fontMap, FONT_ITALIC, font.&setItalic)
                setBooleanFont(fontMap, FONT_STRIKEOUT, font.&setStrikeout)
                if (fontMap.containsKey(FONT_UNDERLINE)) {
                    Object underlineOption = fontMap[FONT_UNDERLINE]
                    byte underline
                    if (underlineOption instanceof Byte) {
                        underline = (byte)underlineOption
                    } else if (underlineOption instanceof Boolean) {
                        underline = FontType.U_SINGLE
                    } else if (underlineOption instanceof String) {
                        switch (underlineOption) {
                            case 'single':
                                underline = FontType.U_SINGLE
                                break
                            case 'singleAccounting':
                                underline = FontType.U_SINGLE_ACCOUNTING
                                break
                            case 'double':
                                underline = FontType.U_DOUBLE
                                break
                            case 'doubleAccounting':
                                underline = FontType.U_DOUBLE_ACCOUNTING
                                break
                            default:
                                throw new IllegalArgumentException("[font: [${FONT_UNDERLINE}: ${fontMap[FONT_UNDERLINE]}]] is not a supported value")
                        }
                    } else {
                        throw new IllegalArgumentException("[font: [${FONT_UNDERLINE}: <>]] must be a boolean or string")
                    }
                    font.setUnderline(underline)
                }
                if (fontMap.containsKey(FONT_COLOR)) {
                    font.setColor(getColor(fontMap[FONT_COLOR]))
                }
                if (fontMap.containsKey(FONT_SIZE)) {
                    font.setFontHeight((Double)fontMap[FONT_SIZE])
                }
                if (fontMap.containsKey(FONT_NAME)) {
                    font.setFontName((String)fontMap[FONT_NAME])
                }
            } else {
                throw new IllegalArgumentException('The font option must be an instance of a Map')
            }
            workbookCache.putFont(fontOptions, font)
        }

        cellStyle.setFont(workbookCache.getFont(fontOptions))
    }

    private XSSFColor getColor(Object obj) {
        Color color
        if (obj instanceof Color) {
            color = (Color)obj
        } else if (obj instanceof String) {
            String hex = (String)obj
            if (hex.startsWith('#')) {
                color = Color.decode(hex)
            } else {
                color = Color.decode("#$obj")
            }
        } else {
            throw new IllegalArgumentException("${obj} must be an instance of ${Color.canonicalName} or a hex ${String.canonicalName}")
        }
        new XSSFColor(color, workbookCache.colorMap)
    }

    @SuppressWarnings('UnnecessaryGetter')
    private BorderStyle getBorderStyle(Object obj) {
        if (obj instanceof BorderStyle) {
            return (BorderStyle)obj
        }

        throw new IllegalArgumentException("The border style must be an instance of ${BorderStyle.getCanonicalName()}")
    }

    private void setBorder(Map border, BorderSide side, BorderStyleApplier styleApplier) {
        final String KEY = side.name().toLowerCase()
        if (border.containsKey(KEY)) {
            if (border[KEY] instanceof Map) {
                Map edge = (Map) border[KEY]
                if (edge.containsKey(COLOR)) {
                    styleApplier.applyColor(side, getColor(edge[COLOR]))
                }
                if (edge.containsKey(STYLE)) {
                    styleApplier.applyStyle(side, getBorderStyle(edge[STYLE]))
                }
            } else {
                styleApplier.applyStyle(side, getBorderStyle(border[KEY]))
            }
        }
    }

    @SuppressWarnings('UnnecessaryGetter')
    private void setHorizontalAlignment(XSSFCellStyle cellStyle, Object horizontalAlignment) {
        HorizontalAlignment hAlign
        if (horizontalAlignment instanceof HorizontalAlignment) {
            hAlign = (HorizontalAlignment)horizontalAlignment
        } else if (horizontalAlignment instanceof String) {
            hAlign = HorizontalAlignment.valueOf(((String)horizontalAlignment).toUpperCase())
        }

        if (hAlign != null) {
            cellStyle.setAlignment(hAlign)
        } else {
            throw new IllegalArgumentException("The horizontal alignment must be an instance of ${HorizontalAlignment.getCanonicalName()}")
        }
    }

    @SuppressWarnings('UnnecessaryGetter')
    private void setVerticalAlignment(XSSFCellStyle cellStyle, Object verticalAlignment) {
        VerticalAlignment vAlign
        if (verticalAlignment instanceof VerticalAlignment) {
            vAlign = (VerticalAlignment) verticalAlignment
        } else if (verticalAlignment instanceof String) {
            vAlign = VerticalAlignment.valueOf(((String)verticalAlignment).toUpperCase())
        }

        if (vAlign != null) {
            cellStyle.setVerticalAlignment(vAlign)
        } else {
            throw new IllegalArgumentException("The vertical alignment must be an instance of ${VerticalAlignment.getCanonicalName()}")
        }
    }

    private void setWrapped(XSSFCellStyle cellStyle, Object wrapped) {
        if (wrapped instanceof Boolean) {
            cellStyle.setWrapText((Boolean)wrapped)
        } else {
            throw new IllegalArgumentException("The wrapped option must be an instance of ${Boolean.canonicalName}")
        }
    }

    private void setLocked(XSSFCellStyle cellStyle, Object locked) {
        if (locked instanceof Boolean) {
            cellStyle.setLocked((Boolean)locked)
        } else {
            throw new IllegalArgumentException("The wrapped option must be an instance of ${Boolean.canonicalName}")
        }
    }

    private void setHidden(XSSFCellStyle cellStyle, Object hidden) {
        if (hidden instanceof Boolean) {
            cellStyle.setHidden((Boolean)hidden)
        } else {
            throw new IllegalArgumentException("The wrapped option must be an instance of ${Boolean.canonicalName}")
        }
    }

    private void setBorder(BorderStyleApplier styleApplier, Map border) {
        if (border.containsKey(STYLE)) {
            BorderStyle style = getBorderStyle(border[STYLE])
            styleApplier.applyStyle(style)
        }
        if (border.containsKey(COLOR)) {
            XSSFColor color = getColor(border[COLOR])
            styleApplier.applyColor(color)
        }
        setBorder(border, BorderSide.LEFT, styleApplier)
        setBorder(border, BorderSide.RIGHT, styleApplier)
        setBorder(border, BorderSide.BOTTOM, styleApplier)
        setBorder(border, BorderSide.TOP, styleApplier)
    }

    private void setFill(XSSFCellStyle cellStyle, Object fill) {
        FillPatternType fillPattern
        if (fill instanceof FillPatternType) {
            fillPattern = (FillPatternType) fill
        } else if (fill instanceof String) {
            fillPattern = FillPatternType.valueOf(((String)fill).toUpperCase())
        }

        if (fillPattern != null) {
            cellStyle.setFillPattern(fillPattern)
        } else {
            throw new IllegalArgumentException("The fill pattern must be an instance of ${Short.canonicalName}")
        }
    }

    private void setForegroundColor(XSSFCellStyle cellStyle, Object foregroundColor) {
        if (cellStyle.fillPatternEnum == FillPatternType.NO_FILL) {
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        }
        cellStyle.setFillForegroundColor(getColor(foregroundColor))
    }

    private void setBackgroundColor(XSSFCellStyle cellStyle, Object backgroundColor) {
        if (cellStyle.fillPatternEnum == FillPatternType.NO_FILL) {
            setForegroundColor(cellStyle, backgroundColor)
        } else {
            cellStyle.setFillBackgroundColor(getColor(backgroundColor))
        }
    }

    /**
     * A method to build a style object based on a map of parameters
     *
     * @param value The data being rendered with the style
     * @param options A map of options to configure the style
     * @return A cell style to apply
     */
    XSSFCellStyle buildStyle(Object value, Map options) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle()
        if (options.containsKey(FORMAT)) {
            setFormat(cellStyle, options[FORMAT])
        }
        if (options.containsKey(FONT)) {
            setFont(cellStyle, options[FONT])
        }
        if (options.containsKey(HIDDEN)) {
            setHidden(cellStyle, options[HIDDEN])
        }
        if (options.containsKey(LOCKED)) {
            setLocked(cellStyle, options[LOCKED])
        }
        if (options.containsKey(WRAPPED)) {
            setWrapped(cellStyle, options[WRAPPED])
        }
        if (options.containsKey(HORIZONTAL_ALIGNMENT)) {
            setHorizontalAlignment(cellStyle, options[HORIZONTAL_ALIGNMENT])
        }
        if (options.containsKey(VERTICAL_ALIGNMENT)) {
            setVerticalAlignment(cellStyle, options[VERTICAL_ALIGNMENT])
        }
        if (options.containsKey(ROTATION)) {
            cellStyle.setRotation((short) options[ROTATION])
        }
        if (options.containsKey(INDENT)) {
            cellStyle.setIndention((short) options[INDENT])
        }
        if (options.containsKey(BORDER)) {
            setBorder(new CellStyleBorderStyleApplier(cellStyle), (Map)options[BORDER])
        }
        if (options.containsKey(FILL)) {
            setFill(cellStyle, options[FILL])
        }
        if (options.containsKey(FOREGROUND_COLOR)) {
            setForegroundColor(cellStyle, options[FOREGROUND_COLOR])
        }
        if (options.containsKey(BACKGROUND_COLOR)) {
            setBackgroundColor(cellStyle, options[BACKGROUND_COLOR])
        }
        cellStyle
    }

    private XSSFCellStyle getStyle(Object value, Map options, Map defaultOptions = null) {
        convertSimpleOptions(options)
        convertSimpleOptions(defaultOptions)
        options = merge(defaultOptions, options)
        if (!options.containsKey(FORMAT) && value != null) {
            Object format = Excel.getFormat(value.class)
            if (format != null) {
                options.put(FORMAT, format)
            }
        }
        if (options) {
            if (workbookCache.containsStyle(options)) {
                workbookCache.getStyle(options)
            } else {
                XSSFCellStyle style = buildStyle(value, options)
                workbookCache.putStyle(options, style)
                style
            }
        } else {
            null
        }
    }

    /**
     * A method to set a style to a cell based on a map of options and a map of default options
     *
     * @param value The data to be rendered to the cell
     * @param cell The cell to apply the styling to
     * @param _options A map of options for styling
     * @param defaultOptions A map of default options for styling
     */
     void setStyle(Object value, SXSSFCell cell, Map options, Map defaultOptions = null) {
         XSSFCellStyle cellStyle = getStyle(value, options, defaultOptions)
         if (cellStyle != null) {
             cell.cellStyle = cellStyle
         }
    }

    /**
     * A method to set a style to a row based on a map of options and a map of default options
     *
     * @param row The row to apply the styling to
     * @param _options A map of options for styling
     * @param defaultOptions A map of default options for styling
     */
    void setStyle(SXSSFRow row, Map options, Map defaultOptions = null) {
        XSSFCellStyle cellStyle = getStyle(null, options, defaultOptions)
        if (cellStyle != null) {
            row.setRowStyle(cellStyle)
        }
    }

    /**
     * Merges multiple maps
     *
     * @param sources The maps to merge
     * @return The merged map
     */
    @CompileStatic(TypeCheckingMode.SKIP)
    Map merge(Map[] sources) {
        if (sources.length == 0) {
            return [:]
        }
        if (sources.length == 1) {
            return sources[0]
        }

        (Map)sources.inject([:]) { result, source ->
            source.each { k, v ->
                result[k] = result[k] instanceof Map ? merge((Map)result[k], (Map)v) : v
            }
            result
        }
    }

    void applyBorderToRegion(CellRangeBorderStyleApplier borderStyleApplier, Map border) {
        setBorder(borderStyleApplier, border)
        borderStyleApplier.setStyles()
    }
    
}
