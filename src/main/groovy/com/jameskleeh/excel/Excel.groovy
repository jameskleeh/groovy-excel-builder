package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.BuiltinFormats

import java.util.concurrent.atomic.AtomicInteger

/**
 * A class to store cell renderers and formatters used to set cell
 * data and format
 */
@CompileStatic
class Excel {

    protected static SortedSet<Entry> rendererEntries = [] as SortedSet
    protected static SortedSet<FormatEntry> formatEntries = [] as SortedSet

    protected static final AtomicInteger RENDERER_SEQUENCE = new AtomicInteger(0)
    protected static final AtomicInteger FORMAT_SEQUENCE = new AtomicInteger(0)

    static {
        registerCellFormat(BigDecimal, 8)
        registerCellFormat(Double, 4)
        registerCellFormat(Float, 4)
        registerCellFormat(Integer, 3)
        registerCellFormat(Long, 3)
        registerCellFormat(Short, 3)
        registerCellFormat(BigInteger, 3)
        registerCellFormat(Date, 'm/d/yyyy')
    }

    static void registerCellRenderer(Class clazz, Integer priority, Closure callable) {
        rendererEntries.add(new Entry(clazz, callable, priority))
    }

    static void registerCellRenderer(Class clazz, Closure callable) {
        registerCellRenderer(clazz, -1, callable)
    }

    static void registerCellFormat(Class clazz, Integer priority, String format) {
        int builtInFormat = BuiltinFormats.getBuiltinFormat(format)
        if (builtInFormat > -1) {
            registerCellFormat(clazz, priority, builtInFormat)
        } else {
            formatEntries.add(new FormatEntry(clazz, format, priority))
        }
    }

    static void registerCellFormat(Class clazz, String format) {
        registerCellFormat(clazz, -1, format)
    }

    static void registerCellFormat(Class clazz, Integer priority, int format) {
        formatEntries.add(new FormatEntry(clazz, format, priority))
    }

    static void registerCellFormat(Class clazz, int format) {
        registerCellFormat(clazz, -1, format)
    }

    static Closure getRenderer(Class clazz) {
        for (Entry entry : rendererEntries) {
            if (entry.clazz == clazz || entry.clazz.isAssignableFrom(clazz)) {
                return entry.renderer
            }
        }
        null
    }

    static Object getFormat(Class clazz) {
        for (FormatEntry entry : formatEntries) {
            if (entry.clazz == clazz || entry.clazz.isAssignableFrom(clazz)) {
                return entry.format
            }
        }
        null
    }

    private static class Entry implements Comparable<Entry> {
        protected final Closure renderer
        protected final Class clazz
        private final int priority
        private final int seq

        Entry(Class clazz, Closure renderer, int priority) {
            this.clazz = clazz
            this.renderer = renderer
            this.priority = priority
            seq = RENDERER_SEQUENCE.incrementAndGet()
        }

        int compareTo(Entry entry) {
            priority == entry.priority ? entry.seq - seq : entry.priority - priority
        }
    }

    private static class FormatEntry implements Comparable<FormatEntry> {
        protected final Object format
        protected final Class clazz
        private final int priority
        private final int seq

        FormatEntry(Class clazz, Object format, int priority) {
            this.clazz = clazz
            this.format = format
            this.priority = priority
            seq = FORMAT_SEQUENCE.incrementAndGet()
        }

        int compareTo(FormatEntry entry) {
            priority == entry.priority ? entry.seq - seq : entry.priority - priority
        }
    }
}
