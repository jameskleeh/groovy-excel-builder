package com.jameskleeh.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Shared
import spock.lang.Specification

/**
 * Created by jameskleeh on 7/3/17.
 */
class WorkbookCacheSpec extends Specification {

    @Shared XSSFWorkbook workbook

    void setupSpec() {
        workbook = new XSSFWorkbook()
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
