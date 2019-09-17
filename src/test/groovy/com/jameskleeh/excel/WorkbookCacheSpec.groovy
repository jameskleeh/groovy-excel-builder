package com.jameskleeh.excel

import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Shared
import spock.lang.Specification

/**
 * Created by jameskleeh on 7/3/17.
 */
class WorkbookCacheSpec extends Specification {

    @Shared SXSSFWorkbook workbook

    void setupSpec() {
        workbook = new SXSSFWorkbook()
    }

    void "test contains style"() {
        given:
        WorkbookCache cache = new WorkbookCache(workbook)

        when:
        cache.putStyle([foo: 'bar'], workbook.createCellStyle())

        then:
        cache.containsStyle([foo: 'bar'])

        when:
        cache.putStyle([foo: 'bar', x: [value: 1]], workbook.createCellStyle())

        then:
        cache.containsStyle([foo: 'bar', x: [value: 1]])

        expect:
        !cache.containsStyle([foo: 'bar', x: [value: 2]])
        cache.styles.size() == 2
    }

}
