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

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.BuiltinFormats

import java.util.concurrent.atomic.AtomicInteger

/**
 * A class to store cell renderers and formatters used to set cell
 * data and format
 *
 * @author James Kleeh
 * @since 0.1.0
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

    /**
     * Registers a cell renderer. What is returned from the closure will be set to the cell value
     *
     * @param clazz The class to render
     * @param priority The priority
     * @param callable The closure to call
     */
    static void registerCellRenderer(Class clazz, Integer priority, Closure callable) {
        rendererEntries.add(new Entry(clazz, callable, priority))
    }

    /**
     * Registers a cell renderer. What is returned from the closure will be set to the cell value
     *
     * @param clazz The class to render
     * @param callable The closure to call
     */
    static void registerCellRenderer(Class clazz, Closure callable) {
        registerCellRenderer(clazz, -1, callable)
    }

    /**
     * Registers a cell format
     *
     * @param clazz The class to render
     * @param priority The priority
     * @param format The format to apply
     */
    static void registerCellFormat(Class clazz, Integer priority, String format) {
        int builtInFormat = BuiltinFormats.getBuiltinFormat(format)
        if (builtInFormat > -1) {
            registerCellFormat(clazz, priority, builtInFormat)
        } else {
            formatEntries.add(new FormatEntry(clazz, format, priority))
        }
    }

    /**
     * Registers a cell format
     *
     * @param clazz The class to render
     * @param format The format to apply
     */
    static void registerCellFormat(Class clazz, String format) {
        registerCellFormat(clazz, -1, format)
    }

    /**
     * Registers a cell format
     *
     * @param clazz The class to render
     * @param priority The priority
     * @param format The format to apply
     */
    static void registerCellFormat(Class clazz, Integer priority, int format) {
        formatEntries.add(new FormatEntry(clazz, format, priority))
    }

    /**
     * Registers a cell format
     *
     * @param clazz The class to render
     * @param format The format to apply
     */
    static void registerCellFormat(Class clazz, int format) {
        registerCellFormat(clazz, -1, format)
    }

    /**
     * Queries the renderers registered
     *
     * @param clazz The class to search for
     * @return The renderer. Null if not found
     */
    static Closure getRenderer(Class clazz) {
        for (Entry entry : rendererEntries) {
            if (entry.clazz == clazz || entry.clazz.isAssignableFrom(clazz)) {
                return entry.renderer
            }
        }
        null
    }

    /**
     * Queries the formats registered
     *
     * @param clazz The class to search for
     * @return The format. Null if not found
     */
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
