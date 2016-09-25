package com.jameskleeh.excel

import groovy.transform.CompileStatic
import groovy.transform.TypeCheckingMode
import org.apache.poi.ss.usermodel.Font as FontType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide
import java.awt.Color

/**
 * A class to build an {@link org.apache.poi.xssf.usermodel.XSSFCellStyle} from a map
 *
 * @author James Kleeh
 * @since 1.0.0
 */
@CompileStatic
class CellStyleBuilder {

    XSSFWorkbook workbook

    private static final Map<XSSFWorkbook, WorkbookCache> WORKBOOK_CACHE = [:]
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
    protected static final String FILL = 'fill'
    protected static final String BACKGROUND_COLOR = 'backgroundColor'
    protected static final String FOREGROUND_COLOR = 'foregroundColor'

    CellStyleBuilder(XSSFWorkbook workbook) {
        this.workbook = workbook
        if (!WORKBOOK_CACHE.containsKey(workbook)) {
            WORKBOOK_CACHE.put(workbook, new WorkbookCache(workbook))
        }
    }

    private static void convertBorderOptions(Map options, String key) {
        if (options.containsKey(key) && options[key] instanceof Border) {
            Border border = (Border)options.remove(key)
            options.put(key, [style: border])
        }
    }

    /**
     *
     * A method to convert global options into specific options.
     * Example:
     * [border: Border.THIN] would convert to
     * [border: [style: Border.THIN, left: [style: Border.THIN], right: ...]]
     *
     * @param options A map of options
     */
     static void convertSimpleOptions(Map options) {
        if (options) {
            if (options.containsKey(BORDER) && options[BORDER] instanceof Border) {
                Border border = (Border)options.remove(BORDER)
                options.put(BORDER, [style: border])
            }
            if (options.containsKey(FONT) && options[FONT] instanceof Font) {
                Font font = (Font)options.remove(FONT)
                Map fontOptions = [:]
                switch (font) {
                    case Font.BOLD:
                        fontOptions[FONT_BOLD] = true
                        break
                    case Font.ITALIC:
                        fontOptions[FONT_ITALIC] = true
                        break
                    case Font.STRIKEOUT:
                        fontOptions[FONT_STRIKEOUT] = true
                        break
                    case Font.UNDERLINE:
                        fontOptions[FONT_UNDERLINE] = (byte)1
                        break
                }
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
        if (format instanceof Short) {
            cellStyle.setDataFormat(format)
        } else if (format instanceof String) {
            cellStyle.setDataFormat(workbook.creationHelper.createDataFormat().getFormat(format))
        } else {
            throw new IllegalArgumentException('The cell format must be a short or String')
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
        WorkbookCache workbookCache = WORKBOOK_CACHE.get(workbook)

        if (!workbookCache.containsFont(fontOptions)) {
            XSSFFont font = workbook.createFont()
            if (fontOptions instanceof Map) {
                Map fontMap = (Map)fontOptions
                setBooleanFont(fontMap, FONT_BOLD, font.&setBold)
                setBooleanFont(fontMap, FONT_ITALIC, font.&setItalic)
                setBooleanFont(fontMap, FONT_STRIKEOUT, font.&setStrikeout)
                if (fontMap.containsKey(FONT_UNDERLINE)) {
                    byte underline
                    if (fontMap[FONT_UNDERLINE] instanceof Boolean) {
                        underline = FontType.U_SINGLE
                    } else if (fontMap[FONT_UNDERLINE] instanceof String) {
                        switch (fontMap[FONT_UNDERLINE]) {
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
                color = Color.decode(obj)
            } else {
                color = Color.decode("#$obj")
            }
        } else {
            throw new IllegalArgumentException("${obj} must be an instance of ${Color.canonicalName} or a hex ${String.canonicalName}")
        }
        new XSSFColor(color)
    }

    private short getBorderStyle(Object obj) {
        if (obj instanceof Border) {
            return (short)obj.ordinal()
        }
        throw new IllegalArgumentException("The border style must be an instance of ${Border.canonicalName}")
    }

    private void setBorder(Map border, String key, Closure borderCallable, Closure colorCallable) {
        if (border.containsKey(key)) {
            if (border[key] instanceof Map) {
                Map edge = (Map) border[key]
                if (edge.containsKey(COLOR)) {
                    colorCallable.call(getColor(edge[COLOR]))
                }
                if (edge.containsKey(STYLE)) {
                    borderCallable.call(getBorderStyle(edge[STYLE]))
                }
            } else {
                borderCallable.call(getBorderStyle(border[key]))
            }
        }
    }

    private void setHorizontalAlignment(XSSFCellStyle cellStyle, Object horizontalAlignment) {
        HorizontalAlignment hAlign
        if (horizontalAlignment instanceof HorizontalAlignment) {
            hAlign = (HorizontalAlignment)horizontalAlignment
        } else if (horizontalAlignment instanceof String) {
            hAlign = HorizontalAlignment.valueOf(horizontalAlignment.toUpperCase())
        }

        if (hAlign != null) {
            cellStyle.setAlignment((short)hAlign.ordinal())
        } else {
            throw new IllegalArgumentException("The horizontal alignment must be an instance of ${HorizontalAlignment.canonicalName}")
        }
    }

    private void setVerticalAlignment(XSSFCellStyle cellStyle, Object verticalAlignment) {
        VerticalAlignment vAlign
        if (verticalAlignment instanceof VerticalAlignment) {
            vAlign = (VerticalAlignment) verticalAlignment
        } else if (verticalAlignment instanceof String) {
            vAlign = VerticalAlignment.valueOf(verticalAlignment.toUpperCase())
        }

        if (vAlign != null) {
            cellStyle.setVerticalAlignment((short)vAlign.ordinal())
        } else {
            throw new IllegalArgumentException("The vertical alignment must be an instance of ${VerticalAlignment.canonicalName}")
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

    private void setBorder(XSSFCellStyle cellStyle, Map border) {
        if (border.containsKey(STYLE)) {
            short style = getBorderStyle(border[STYLE])
            cellStyle.setBorderBottom(style)
            cellStyle.setBorderTop(style)
            cellStyle.setBorderLeft(style)
            cellStyle.setBorderRight(style)
        }
        if (border.containsKey(COLOR)) {
            XSSFColor color = getColor(border[COLOR])
            cellStyle.setBorderColor(BorderSide.BOTTOM, color)
            cellStyle.setBorderColor(BorderSide.TOP, color)
            cellStyle.setBorderColor(BorderSide.LEFT, color)
            cellStyle.setBorderColor(BorderSide.RIGHT, color)
        }
        setBorder(border, LEFT, cellStyle.&setBorderLeft, cellStyle.&setLeftBorderColor)
        setBorder(border, RIGHT, cellStyle.&setBorderRight, cellStyle.&setRightBorderColor)
        setBorder(border, BOTTOM, cellStyle.&setBorderBottom, cellStyle.&setBottomBorderColor)
        setBorder(border, TOP, cellStyle.&setBorderTop, cellStyle.&setTopBorderColor)
    }

    private void setFill(XSSFCellStyle cellStyle, Object fill) {
        if (fill instanceof Fill) {
            cellStyle.setFillPattern((short)((Fill)fill).ordinal())
        } else {
            throw new IllegalArgumentException("The fill pattern must be an instance of ${Short.canonicalName}")
        }
    }

    private void setForegroundColor(XSSFCellStyle cellStyle, Object foregroundColor) {
        cellStyle.setFillForegroundColor(getColor(foregroundColor))
    }

    private void setBackgroundColor(XSSFCellStyle cellStyle, Object backgroundColor) {
        XSSFColor color = getColor(backgroundColor)
        if (cellStyle.fillForegroundColor == IndexedColors.AUTOMATIC.index) {
            cellStyle.setFillForegroundColor(color)
        } else {
            cellStyle.setFillBackgroundColor(color)
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
        XSSFCellStyle cellStyle = workbook.createCellStyle()
        if (options.containsKey(FORMAT)) {
            setFormat(cellStyle, options[FORMAT])
        } else {
            Object format = Excel.getFormat(value.class)
            if (format) {
                setFormat(cellStyle, format)
            }
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
            setBorder(cellStyle, (Map)options[BORDER])
        }
        if (options.containsKey(FILL)) {
            setFill(cellStyle, options[FILL])
        } else {
            setFill(cellStyle, Fill.SOLID_FOREGROUND)
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
        options = merge(defaultOptions, options)
        if (options) {
            WorkbookCache workbookCache = WORKBOOK_CACHE.get(workbook)
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
     void setStyle(Object value, XSSFCell cell, Map options, Map defaultOptions = null) {
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
    void setStyle(XSSFRow row, Map options, Map defaultOptions = null) {
        XSSFCellStyle cellStyle = getStyle(null, options, defaultOptions)
        if (cellStyle != null) {
            row.setRowStyle(cellStyle)
        }
    }

    @CompileStatic(TypeCheckingMode.SKIP)
    private Map merge(Map[] sources) {
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
    
}
