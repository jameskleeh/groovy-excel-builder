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
package com.jameskleeh.excel.extensions

import java.util.regex.Matcher
import java.util.regex.Pattern

/**
 * A class to create anchored column references
 *
 * @author James Kleeh
 * @since 0.1.0
 */
class StringExtension {

    static final Pattern DIGIT = Pattern.compile('\\d')

    static String anchorColumn(final String self) {
        '$' + self
    }

    static String anchorRow(final String self) {
        Matcher m = DIGIT.matcher(self)
        int position = 0
        if (m.find()) {
            position = m.start()
        }
        self[0..position - 1] + '$' + self[position.. - 1]
    }

    static String anchor(final String self) {
        self.anchorRow().anchorColumn()
    }
}
