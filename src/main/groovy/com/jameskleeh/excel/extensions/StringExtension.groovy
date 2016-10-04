package com.jameskleeh.excel.extensions

/**
 * A class to create anchored column references
 *
 * @author James Kleeh
 * @since 0.1.0
 */
class StringExtension {

    static String anchorColumn(final String self) {
        '$' + self
    }

    static String anchorRow(final String self) {
        self[0] + '$' + self[1..-1]
    }

    static String anchor(final String self) {
        self.anchorRow().anchorColumn()
    }
}
