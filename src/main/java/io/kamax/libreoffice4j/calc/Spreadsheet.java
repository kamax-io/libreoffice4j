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

import com.sun.star.sheet.*;
import com.sun.star.table.CellAddress;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;
import com.sun.star.uno.UnoRuntime;
import io.kamax.libreoffice4j.exception.IndexOutOfBoundsException;
import io.kamax.libreoffice4j.table.Cell;
import io.kamax.libreoffice4j.table.IRow;
import io.kamax.libreoffice4j.table.Row;
import io.kamax.libreoffice4j.table.RowSerie;

public class Spreadsheet implements ISpreadsheet {

    private XSpreadsheet sheet;

    public Spreadsheet(XSpreadsheet sheet) {
        this.sheet = sheet;
    }

    public Cell getCell(int col, int row) {
        try {
            return new Cell(sheet.getCellByPosition(col, row));
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(col, row);
        }
    }

    public IRow getRow(int row) {
        return new Row(row, sheet);
    }

    public RowSerie getRowSeries(int startIndex) {
        return new RowSerie(startIndex, sheet);
    }

    protected CellRangeAddress getRowAddress(int row) {
        try {
            return UnoRuntime.queryInterface(XCellRangeAddressable.class, sheet.getCellRangeByPosition(0, row, 150, row)).getRangeAddress();
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(0, row);
        }
    }

    protected CellAddress getCellAddress(XSpreadsheet sheet, int column, int row) {
        try {
            return UnoRuntime.queryInterface(XCellAddressable.class, sheet.getCellByPosition(column, row)).getCellAddress();
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(column, row);
        }
    }

    public void insertRowAboveAndCopy(int row) {
        XCellRangeMovement mover = UnoRuntime.queryInterface(XCellRangeMovement.class, sheet);
        mover.insertCells(getRowAddress(row), CellInsertMode.ROWS);
        mover.copyRange(getCellAddress(sheet, 0, row), getRowAddress(row - 1));
    }

    @Override
    public void copyCellRangeToPosition(String sourceRange, String destinationCell) {
        XCellRange oRange = UnoRuntime.queryInterface(XCellRange.class, sheet).getCellRangeByName(sourceRange);
        CellAddress destAddress = UnoRuntime.queryInterface(XCellAddressable.class, sheet.getCellRangeByName(destinationCell)).getCellAddress();
        CellRangeAddress rangeAddres = UnoRuntime.queryInterface(XCellRangeAddressable.class, oRange).getRangeAddress();
        UnoRuntime.queryInterface(XCellRangeMovement.class, sheet).copyRange(destAddress, rangeAddres);
    }

}
