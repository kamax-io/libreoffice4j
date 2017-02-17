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

package io.kamax.libreoffice4j;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.io.IOException;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import io.kamax.libreoffice4j.calc.SpreadsheetDocument;
import io.kamax.libreoffice4j.exception.LibreOfficeException;

import java.io.File;

public class LibreOfficeManager {

    private static LibreOfficeManager mgr;

    private static XComponentLoader initContext() {
        try {
            XComponentContext xContext = Bootstrap.bootstrap();
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
            return UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);
        } catch (Exception e) {
            throw new LibreOfficeException(e);
        }
    }

    public static synchronized LibreOfficeManager getInstance() {
        try {
            if (mgr == null) {
                XComponentLoader xCLoader = initContext();
                mgr = new LibreOfficeManager(xCLoader);
            }

            return mgr;
        } catch (Exception e) {
            throw new LibreOfficeException(e);
        }
    }

    private XComponentLoader xCLoader;

    private LibreOfficeManager(XComponentLoader xCLoader) {
        this.xCLoader = xCLoader;
    }

    public SpreadsheetDocument loadSpreadsheetDocument(File file) {
        try {
            PropertyValue[] props = new PropertyValue[1];
            props[0] = new PropertyValue();
            props[0].Name = "Hidden";
            props[0].Value = true;

            XComponent xComp = xCLoader.loadComponentFromURL("file://" + file.getAbsolutePath(), "_blank", 0, props);
            XSpreadsheetDocument doc = UnoRuntime.queryInterface(XSpreadsheetDocument.class, xComp);

            return new SpreadsheetDocument(doc);
        } catch (DisposedException e) {
            xCLoader = initContext();
            return loadSpreadsheetDocument(file);
        } catch (IOException e) {
            throw new LibreOfficeException(e);
        }
    }

}
