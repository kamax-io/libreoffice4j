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

package io.kamax.libreoffice4j.calc;

import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.sheet.XSpreadsheet;
import io.kamax.libreoffice4j.table.Cell;
import io.kamax.libreoffice4j.table.RowSerie;

public class Spreadsheet implements ISpreadsheet {

    private XSpreadsheet sheet;

    public Spreadsheet(XSpreadsheet sheet) {
        this.sheet = sheet;
    }

    public Cell getCell(int col, int row) {
        try {
            return new Cell(sheet.getCellByPosition(col, row));
        } catch (IndexOutOfBoundsException e) {
            throw new ArrayIndexOutOfBoundsException(col + ":" + row);
        }
    }

    public RowSerie getRowSeries(int startIndex) {
        return new RowSerie(startIndex, sheet);
    }

}
