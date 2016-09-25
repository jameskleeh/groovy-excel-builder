package com.jameskleeh.excel.extensions

/**
 * Created by jameskleeh on 9/25/16.
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
