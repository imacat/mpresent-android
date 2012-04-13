/*
 * ImpressController.java
 *
 * Created on 2012-01-06
 * 
 * Copyright (c) 2012 imacat
 */

/*
 * Copyright (c) 2012 imacat
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package tw.idv.imacat.android.mpresent.uno;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.ResourceBundle;

import com.sun.star.awt.XBitmap;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XSlideRenderer;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.presentation.XPresentation;
import com.sun.star.presentation.XPresentation2;
import com.sun.star.presentation.XPresentationSupplier;
import com.sun.star.presentation.XSlideShowController;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.ucb.XSimpleFileAccess3;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * The OpenOffice.org Impress Presentation Controller.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class ImpressController {
    
    /** The OpenOffice.org remote connection. */
    private OfficeConnection conn = null;
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The file name converter. */
    private XFileIdentifierConverter fileNameConverter = null;
    
    /** The slide renderer. */
    private XSlideRenderer slideRenderer = null;
    
    /** The loaded presentation documents. */
    private ArrayList<OfficeConnection.Document> loaded
        = new ArrayList<OfficeConnection.Document>();
    
    /** The running presentation document. */
    private OfficeConnection.Document running = null;
    
    /** The index of the running presentation in the loaded presentations. */
    private int runningIndex = -1;
    
    /** The draw pages of the running presentation. */
    private XDrawPages pages = null;
    
    /** The index of the current slide of the running presentation. */
    private int curSlide = -1;
    
    /** The number of slides of the running presentation. */
    private int slideCount = -1;
    
    /**
     * Creates a new instance of an OpenOffice.org Impress Presentation
     * Controller for Android.
     *
     * @param conn the host name
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     */
    protected ImpressController(OfficeConnection conn) {
        this.conn = conn;
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.fileNameConverter = this.conn.getFileNameConverterService();
        this.slideRenderer = this.conn.getSlideRendererService();
        return;
    }
    
    /**
     * Starts the presentation.
     *
     * @param doc the presentation document to start.
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.ClosedException if
     *         the presentation document is closed
     */
    public void start(OfficeConnection.Document doc)
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.ClosedException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        
        this.checkConnection();
        this.refresh();
        
        try {
            xPresentationSupplier = (XPresentationSupplier)
                UnoRuntime.queryInterface(
                XPresentationSupplier.class, doc.getComponent());
            xPresentation = xPresentationSupplier.getPresentation();
        } catch (com.sun.star.lang.DisposedException e) {
            throw new tw.idv.imacat.android.mpresent.uno.ClosedException(
                this._("err_presentation_closed"));
        }
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        if (!xPresentation2.isRunning()) {
            if (this.running != null) {
                xPresentationSupplier = (XPresentationSupplier)
                    UnoRuntime.queryInterface(
                    XPresentationSupplier.class, this.running.getComponent());
                xPresentation = xPresentationSupplier.getPresentation();
                xPresentation.end();
                
                xPresentationSupplier = (XPresentationSupplier)
                    UnoRuntime.queryInterface(
                    XPresentationSupplier.class, doc.getComponent());
                xPresentation = xPresentationSupplier.getPresentation();
            }
            xPresentation.start();
        }
        return;
    }
    
    /**
     * Stops the presentation.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.NoRunningException if there
     *         is no running presentation.
     */
    public void stop()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.NoRunningException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        
        this.checkConnection();
        this.refresh();
        
        if (this.running == null) {
            throw new tw.idv.imacat.android.mpresent.uno.NoRunningException(
                this._("err_no_running"));
        }
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.running.getComponent());
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation.end();
        return;
    }
    
    /**
     * Goes to the previous-previous slide.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.NoRunningException if there
     *         is no running presentation.
     */
    public void gotoPreviousPreviousSlide()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.NoRunningException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        this.checkConnection();
        this.refresh();
        
        if (this.running == null) {
            throw new tw.idv.imacat.android.mpresent.uno.NoRunningException(
                this._("err_no_running"));
        }
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.running.getComponent());
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        xSlideShowController.gotoPreviousSlide();
        xSlideShowController.gotoPreviousSlide();
        return;
    }
    
    /**
     * Goes to the previous slide.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.NoRunningException if there
     *         is no running presentation.
     */
    public void gotoPreviousSlide()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.NoRunningException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        this.checkConnection();
        this.refresh();
        
        if (this.running == null) {
            throw new tw.idv.imacat.android.mpresent.uno.NoRunningException(
                this._("err_no_running"));
        }
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.running.getComponent());
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        xSlideShowController.gotoPreviousSlide();
        return;
    }
    
    /**
     * Goes to the next slide.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.NoRunningException if there
     *         is no running presentation.
     */
    public void gotoNextSlide()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.NoRunningException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        this.checkConnection();
        this.refresh();
        
        if (this.running == null) {
            throw new tw.idv.imacat.android.mpresent.uno.NoRunningException(
                this._("err_no_running"));
        }
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.running.getComponent());
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        xSlideShowController.gotoNextSlide();
        return;
    }
    
    /**
     * Goes to the next-next slide.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     * @throws tw.idv.imacat.android.mpresent.uno.NoRunningException if there
     *         is no running presentation.
     */
    public void gotoNextNextSlide()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException,
                tw.idv.imacat.android.mpresent.uno.NoRunningException {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        this.checkConnection();
        this.refresh();
        
        if (this.running == null) {
            throw new tw.idv.imacat.android.mpresent.uno.NoRunningException(
                this._("err_no_running"));
        }
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.running.getComponent());
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        xSlideShowController.gotoNextSlide();
        xSlideShowController.gotoNextSlide();
        return;
    }
    
    /**
     * Checks the OpenOffice.org remote connection
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     */
    public void checkConnection()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException {
        if (!this.conn.isConnected()) {
            this.conn.connect();
            this.fileNameConverter = this.conn.getFileNameConverterService();
            this.slideRenderer = this.conn.getSlideRendererService();
        }
        return;
    }
    
    /**
     * Refreshes the remote OpenOffice.org server information.
     *
     */
    public void refresh() {
        List<OfficeConnection.Document> docs
            = this.conn.getLoadedDocuments();
        
        this.loaded = new ArrayList<OfficeConnection.Document>();
        this.running = null;
        this.runningIndex = -1;
        
        // Obtain all the components
        for (int i = 0; i < docs.size(); i++) {
            OfficeConnection.Document doc = docs.get(i);
            XComponent xComponent = doc.getComponent();
            XPresentationSupplier xPresentationSupplier = null;
            XPresentation xPresentation = null;
            XPresentation2 xPresentation2 = null;
            
            if (!doc.isPresentation()) {
                continue;
            }
            this.loaded.add(doc);
            
            xPresentationSupplier = (XPresentationSupplier)
                UnoRuntime.queryInterface(
                XPresentationSupplier.class, xComponent);
            xPresentation = xPresentationSupplier.getPresentation();
            xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
                XPresentation2.class, xPresentation);
            if (xPresentation2.isRunning()) {
                XDrawPagesSupplier xDrawPgaesSupplier = (XDrawPagesSupplier)
                    UnoRuntime.queryInterface(
                    XDrawPagesSupplier.class, xComponent);
                XSlideShowController xSlideShowController
                    = xPresentation2.getController();
                
                this.running = doc;
                this.pages = xDrawPgaesSupplier.getDrawPages();
                this.curSlide = xSlideShowController.getCurrentSlideIndex();
                this.slideCount = xSlideShowController.getSlideCount();
            }
        }
        
        // Removes the duplicated component to the running presentation.
        if (this.running != null) {
            for (int i = 0; i < this.loaded.size(); i++) {
                OfficeConnection.Document doc = this.loaded.get(i);
                
                if (doc.equals(this.running) && doc != this.running) {
                    this.loaded.remove(doc);
                }
            }
            for (int i = 0; i < this.loaded.size(); i++) {
                if (this.loaded.get(i) == this.running) {
                    this.runningIndex = i;
                    break;
                }
            }
        }
        return;
    }
    
    /**
     * Returns the loaded presentation documents.
     *
     * @return the loaded presentation documents
     */
    public ArrayList<OfficeConnection.Document> getLoadedPresentations() {
        return this.loaded;
    }
    
    /**
     * Returns the index of the running presentation in the loaded
     * presentations.
     *
     * @return the index of the running presentation in the loaded
     *         presentations
     */
    public int getRunningIndex() {
        return this.runningIndex;
    }
    
    /**
     * Returns whether a presentation is running.
     *
     * @return true if a presentation is running, false otherwise.
     */
    public boolean isRunning() {
        return this.running != null;
    }
    
    /**
     * Returns the index of the current slide of the running presentation.
     *
     * @return the index of the current slide of the running presentation.
     */
    public int getCurrentSlideIndex() {
        return this.curSlide;
    }
    
    /**
     * Returns the number of slides of the running presentation.
     *
     * @return the number of slides of the running presentation.
     */
    public int getSlideCount() {
        return this.slideCount;
    }
    
    /**
     * Returns the OpenOffice.org remote connection.
     *
     * @return the OpenOffice.org remote connection
     */
    public OfficeConnection getConnection() {
        return this.conn;
    }
    
    /**
     * Create a preview for the given slide that has the same aspect
     * ratio as the page and is as large as possible but not larger
     * than the specified size.  The reason for not using the given
     * size directly as preview size and thus possibly changing the
     * aspect ratio is that a) a different aspect ratio is not used
     * often, and b) leaving the adaption of the actual preview size
     * (according to the aspect ratio of the slide) to the slide
     * renderer is more convenient to the caller than having to this
     * himself. 
     *
     * @param slide the slide for which a preview will be created.
     * @param width the maximum width of the preview measured in pixels.
     * @param height the maximum height of the preview measured in pixels.
     * @return the device independent bitmap (DIB) of the preview.
     */
    public byte[] createPreview(int slide, int width, int height) {
        Object page = null;
        XDrawPage xDrawPage = null;
        com.sun.star.awt.Size size = new com.sun.star.awt.Size();
        XBitmap xBitmap = null;
        
        try {
            page = this.pages.getByIndex(slide);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IllegalArgumentException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
        xDrawPage = (XDrawPage) UnoRuntime.queryInterface(
            XDrawPage.class, page);
        size.Width = width;
        size.Height = height;
        xBitmap = this.slideRenderer.createPreview(
            xDrawPage, size, (short) 1);
        return xBitmap.getDIB();
    }
    
    /**
     * Gets a string for the given key from this resource bundle
     * or one of its parents.  If the key is missing, returns
     * the key itself.
     * 
     * @param key the key for the desired string 
     * @return the string for the given key 
     */
    private String _(String key) {
        try {
            return this.l10n.getString(key);
        } catch (java.util.MissingResourceException e) {
            return key;
        }
    }
}
