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

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.uno.UnoRuntime;
import io.kamax.libreoffice4j.exception.IndexOutOfBoundsException;
import io.kamax.libreoffice4j.exception.LibreOfficeException;

public class Row implements IRow {

    private XSpreadsheet sheet;
    private int index;

    public Row(int index, XSpreadsheet sheet) {
        this.sheet = sheet;
        this.index = index;
    }

    @Override
    public int getIndex() {
        return index;
    }

    @Override
    public Cell getCell(int col) {
        try {
            return new Cell(sheet.getCellByPosition(col, index));
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(col, index);
        }
    }

    @Override
    public void setPageBreak(boolean isPageBreak) {
        try {
            XColumnRowRange tableRows = UnoRuntime.queryInterface(XColumnRowRange.class, sheet);
            XPropertySet tableRow = UnoRuntime.queryInterface(XPropertySet.class, tableRows.getRows().getByIndex(index));
            tableRow.setPropertyValue("IsStartOfNewPage", true);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new IndexOutOfBoundsException(0, index);
        } catch (WrappedTargetException e) {
            throw new LibreOfficeException(e);
        } catch (PropertyVetoException e) {
            throw new LibreOfficeException(e);
        } catch (UnknownPropertyException e) {
            throw new LibreOfficeException(e);
        }
    }

}
