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

import com.sun.star.beans.PropertyValue;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import com.sun.star.view.XPrintable;
import io.kamax.libreoffice4j.exception.LibreOfficeException;

import java.io.File;

public class SpreadsheetDocument implements ISpreadsheetDocument {

    private XSpreadsheetDocument doc;

    public SpreadsheetDocument(XSpreadsheetDocument doc) {
        this.doc = doc;
    }

    public void close() {
        try {
            XCloseable xCloseable = UnoRuntime.queryInterface(XCloseable.class, doc);

            if (xCloseable == null) {
                throw new IllegalStateException("XSpreadsheet is not closable");
            }

            xCloseable.close(false);
        } catch (CloseVetoException e) {
            throw new LibreOfficeException(e);
        }
    }

    public void print() {
        print(1);
    }

    public void print(boolean blocking) {
        print(null, 1, blocking);
    }

    public void print(int amount) {
        print(null, 1, false);
    }

    public void print(String printerName, int amount, boolean blocking) {
        XPrintable xPrintable = UnoRuntime.queryInterface(com.sun.star.view.XPrintable.class, doc);

        if (printerName != null) {
            PropertyValue[] printerProps = new PropertyValue[1];
            printerProps[0] = new PropertyValue();
            printerProps[0].Name = "Name";
            printerProps[0].Value = printerName;
            xPrintable.setPrinter(printerProps);
        }

        PropertyValue[] printJobProps = new PropertyValue[2];
        printJobProps[0] = new PropertyValue();
        printJobProps[0].Name = "CopyCount";
        printJobProps[0].Value = amount;

        printJobProps[1] = new PropertyValue();
        printJobProps[1].Name = "Wait";
        printJobProps[1].Value = blocking;

        xPrintable.print(printJobProps);
    }

    public Spreadsheet getSpreadsheet(int index) {
        try {
            XIndexAccess sheets = UnoRuntime.queryInterface(XIndexAccess.class, doc.getSheets());
            XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, sheets.getByIndex(index));
            return new Spreadsheet(sheet);
        } catch (WrappedTargetException e) {
            throw new LibreOfficeException(e);
        } catch (IndexOutOfBoundsException e) {
            throw new LibreOfficeException(e);
        }
    }

    public File saveTo(File location) {
        try {
            location = location.getAbsoluteFile();
            String targetUrl = "file://" + location.toString();

            XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, doc);
            PropertyValue[] docArgs = new PropertyValue[1];
            docArgs[0] = new PropertyValue();
            docArgs[0].Name = "Overwrite";
            docArgs[0].Value = Boolean.TRUE;
            xStorable.storeAsURL(targetUrl, docArgs);

            return location;
        } catch (IOException e) {
            throw new LibreOfficeException(e);
        }
    }

}
