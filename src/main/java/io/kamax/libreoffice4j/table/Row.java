/*
 * LibreOffice4j - LibreOffice Java API made easy
 * Copyright (C) 2017 Maxime Dor
 *
 * https://max.kamax.io/
 *
 * This file is part of LibreOffice4j
 *
 * LibreOffice4j is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

package io.kamax.libreoffice4j.table;

import com.sun.star.sheet.XSpreadsheet;
import io.kamax.libreoffice4j.exception.IndexOutOfBoundsException;

public class Row implements IRow {

    private XSpreadsheet sheet;
    private int index;

    public Row(int index, XSpreadsheet sheet) {
        this.sheet = sheet;
        this.index = index;
    }

    public Cell getCell(int col) {
        try {
            return new Cell(sheet.getCellByPosition(col, index));
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(col, index);
        }
    }

}
