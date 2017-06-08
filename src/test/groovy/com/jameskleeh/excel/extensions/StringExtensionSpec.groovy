package com.jameskleeh.excel.extensions

import spock.lang.Specification
import spock.lang.Subject

@Subject(StringExtension)
class StringExtensionSpec extends Specification {

    void "test anchor column"() {
        expect:
        'A2'.anchorColumn() == '$A2'
    }

    void "test anchor row"() {
        expect:
        'A2'.anchorRow() == 'A$2'
    }

    void "test anchor"() {
        expect:
        'A2'.anchor() == '$A$2'
    }

}
