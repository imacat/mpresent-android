/*
 * OfficeConnection.java
 *
 * Created on 2012-01-20
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

import java.util.Collections;
import java.util.List;
import java.util.ArrayList;
import java.util.ResourceBundle;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XSlideRenderer;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XTitle;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.ucb.XSimpleFileAccess3;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * The OpenOffice.org remote connection.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class OfficeConnection {
    
    /** The remote host. */
    private String host = null;
    
    /** The remote port. */
    private int port = -1;
    
    /** The desktop service. */
    private XDesktop desktop = null;
    
    /** The bootstrap context. */
    private XComponentContext bootstrapContext = null;
    
    /** The registry service manager. */
    private XMultiComponentFactory serviceManager = null;
    
    /** The file name converter. */
    private XFileIdentifierConverter fileNameConverterService = null;
    
    /** The server file access. */
    private XSimpleFileAccess3 fileAccessService = null;
    
    /** The configuration provider. */
    private XMultiServiceFactory configProvider = null;
    
    /** The slide renderer. */
    private XSlideRenderer slideRendererService = null;
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The Impress controller. */
    private ImpressController impressController = null;
    
    /** The file access. */
    private FileAccess fileAccess = null;
    
    /**
     * Creates a new instance of an OpenOffice.org remote connection.
     *
     * @param host the host name
     * @param port the port number
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     */
    public OfficeConnection(String host, int port)
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException {
        this.host = host;
        this.port = port;
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.connect();
        return;
    }
    
    /**
     * Connects to the OpenOffice.org process.
     *
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     */
    protected void connect()
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException {
        XComponentContext localContext = null;
        XMultiComponentFactory localServiceManager = null;
        Object unoUrlResolver = null;
        XUnoUrlResolver xUnoUrlResolver = null;
        Object bootstrapContext = null;
        
        if (this.isConnected()) {
            return;
        }
        
        // Obtains the local context
        try {
            localContext = Bootstrap.createInitialComponentContext(null);
        } catch (java.lang.Exception e) {
            throw new com.sun.star.comp.helper.BootstrapException(e);
        }
        if (localContext == null) {
            throw new com.sun.star.comp.helper.BootstrapException(
                this._("err_ooo_no_lcc"));
        }
        
        // Obtains the local service manager
        localServiceManager = localContext.getServiceManager();
        
        // Obtains the URL resolver
        try {
            unoUrlResolver = localServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", localContext);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
        xUnoUrlResolver = (XUnoUrlResolver) UnoRuntime.queryInterface(
            XUnoUrlResolver.class, unoUrlResolver);
        
        // Obtains the context
        try {
            bootstrapContext = xUnoUrlResolver.resolve(String.format(
                "uno:socket,host=%s,port=%d;urp;StarOffice.ComponentContext",
                this.host, this.port));
        } catch (com.sun.star.connection.NoConnectException e) {
            throw e;
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            throw e;
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        }
        if (bootstrapContext == null) {
            throw new java.lang.UnsupportedOperationException(
                String.format(this._("err_no_conn"), this.host, this.port));
        }
        this.bootstrapContext = (XComponentContext) UnoRuntime.queryInterface(
            XComponentContext.class, bootstrapContext);
        
        // Obtains the service manager
        this.serviceManager = this.bootstrapContext.getServiceManager();
        return;
    }
    
    /**
     * Returns whether the connection is on and alive.
     *
     * @return true if the connection is alive, false otherwise
     */
    public boolean isConnected() {
        if (this.serviceManager == null) {
            return false;
        }
        try {
            UnoRuntime.queryInterface(
                XPropertySet.class, this.serviceManager);
        } catch (com.sun.star.lang.DisposedException e) {
            this.serviceManager = null;
            return false;
        }
        return true;
    }
    
    /**
     * Opens a document.
     *
     * @param url the URL of the document
     * @return the newly-loaded document
     * @throws java.io.IOException when URL couldn't be found or was corrupt
     */
    public Document open(String url)
            throws java.io.IOException {
        XComponentLoader xComponentLoader = (XComponentLoader)
            UnoRuntime.queryInterface(
            XComponentLoader.class, this.getDesktopService());
        PropertyValue props[] = new PropertyValue[0];
        XComponent doc = null;
        
        try {
            doc = xComponentLoader.loadComponentFromURL(
                url, "_default", 0, props);
        } catch (com.sun.star.io.IOException e) {
            if (url.startsWith("file:///")) {
                XFileIdentifierConverter nameConverter
                    = this.getFileNameConverterService();
                
                url = nameConverter.getSystemPathFromFileURL(url);
            }
            throw new java.io.IOException(
                String.format(_("err_open"), url), e);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            if (url.startsWith("file:///")) {
                XFileIdentifierConverter nameConverter
                    = this.getFileNameConverterService();
                
                url = nameConverter.getSystemPathFromFileURL(url);
            }
            throw new java.io.IOException(
                String.format(_("err_open"), url), e);
        }
        return new Document(doc);
    }
    
    /**
     * Returns the loaded documents.
     *
     * @return the loaded documents
     */
    public List<Document> getLoadedDocuments() {
        XEnumerationAccess xEnumerationAccess
            = this.getDesktopService().getComponents();
        XEnumeration xEnumeration = xEnumerationAccess.createEnumeration();
        List<Document> docs = new ArrayList<Document>();
        
        while (xEnumeration.hasMoreElements()) {
            Object component = null;
            XServiceInfo xServiceInfo = null;
            XComponent xComponent = null;
            
            try {
                component = xEnumeration.nextElement();
            } catch (com.sun.star.container.NoSuchElementException e) {
                throw new java.util.NoSuchElementException(
                    e.getLocalizedMessage());
            } catch (com.sun.star.lang.WrappedTargetException e) {
                throw new java.lang.RuntimeException(e);
            }
            
            xServiceInfo = (XServiceInfo) UnoRuntime.queryInterface(
                    XServiceInfo.class, component);
            if (xServiceInfo == null) {
                continue;
            }
            if (!xServiceInfo.supportsService(
                    "com.sun.star.document.OfficeDocument")) {
                continue;
            }
            
            xComponent = (XComponent) UnoRuntime.queryInterface(
                XComponent.class, component);
            docs.add(new Document(xComponent));
        }
        Collections.sort(docs);
        return docs;
    }
    
    /**
     * Returns the value of a configuration.
     *
     * @param path the path of the confuration
     * @returns the value of the configuration
     */
    protected Object getConfiguration(String path) {
        Object value = null;
        int start = 0, pos = -1;
        PropertyValue args[] = new PropertyValue[1];
        
        if (this.configProvider == null) {
            Object service = null;
            
            try {
                service = this.serviceManager.createInstanceWithContext(
                        "com.sun.star.configuration.ConfigurationProvider",
                        this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            this.configProvider = (XMultiServiceFactory)
                UnoRuntime.queryInterface(
                XMultiServiceFactory.class, service);
        }
        
        args[0] = new PropertyValue();
        args[0].Name = "nodepath";
        pos = path.indexOf('/', 1);
        if (pos == -1) {
            args[0].Value = path;
        } else {
            args[0].Value = path.substring(0, pos);
        }
        try {
            value = this.configProvider.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", args);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.IllegalArgumentException(e);
        }
        
        while (pos != -1) {
            XNameAccess xNameAccess = (XNameAccess)
                UnoRuntime.queryInterface(
                XNameAccess.class, value);
            String name = null;
            
            start = pos + 1;
            pos = path.indexOf('/', start);
            if (pos == -1) {
                name = path.substring(start);
            } else {
                name = path.substring(start, pos);
            }
            try {
                value = xNameAccess.getByName(name);
            } catch (com.sun.star.container.NoSuchElementException e) {
                throw new java.lang.IllegalArgumentException(e);
            } catch (com.sun.star.lang.WrappedTargetException e) {
                throw new java.lang.RuntimeException(e);
            }
        }
        return value;
    }
    
    /**
     * Returns an Impress controller.
     *
     * @return an Impress controller
     */
    public ImpressController getImpressController() {
        if (this.impressController == null) {
            this.impressController = new ImpressController(this);
        }
        return this.impressController;
    }
    
    /**
     * Returns a file access.
     *
     * @return a file access
     */
    public FileAccess getFileAccess() {
        if (this.fileAccess == null) {
            this.fileAccess = new FileAccess(this);
        }
        return this.fileAccess;
    }
    
    /**
     * Returns the stack trace of an exception.
     * <p>
     * This method should not belong here, but only for convienence.
     *
     * @param e the exception
     * @return the stack trace of the exception
     */
    public static String getStackTrace(Exception e) {
        String message = e.getClass().getName()
            + ": " + e.getLocalizedMessage() + "\n";
        StackTraceElement trace[] = e.getStackTrace();
        
        for (int i = 0; i < trace.length; i++) {
            message += "    at " + trace[i].getClassName()
                + "(" + trace[i].getFileName()
                + ":" + trace[i].getLineNumber() + ")\n";
        }
        return message;
    }
    
    /**
     * Returns the file name converter service.
     *
     * @return the file name converter service
     */
    protected XFileIdentifierConverter getFileNameConverterService() {
        if (this.fileNameConverterService == null) {
            Object service = null;
            
            try {
                service = this.serviceManager.createInstanceWithContext(
                        "com.sun.star.ucb.FileContentProvider",
                        this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            this.fileNameConverterService = (XFileIdentifierConverter)
                UnoRuntime.queryInterface(
                XFileIdentifierConverter.class, service);
        }
        return this.fileNameConverterService;
    }
    
    /**
     * Returns the file access service.
     *
     * @return the file access service
     */
    protected XSimpleFileAccess3 getFileAccessService() {
        if (this.fileAccessService == null) {
            Object service = null;
            
            try {
                service = this.serviceManager.createInstanceWithContext(
                        "com.sun.star.ucb.SimpleFileAccess",
                        this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            this.fileAccessService = (XSimpleFileAccess3)
                UnoRuntime.queryInterface(
                XSimpleFileAccess3.class, service);
        }
        return this.fileAccessService;
    }
    
    /**
     * Returns the slide renderer service.
     *
     * @return the slide renderer service
     */
    protected XSlideRenderer getSlideRendererService() {
        //if (this.slideRendererService == null) {
            Object service = null;
            
            try {
                service = this.serviceManager.createInstanceWithContext(
                        "com.sun.star.drawing.SlideRenderer",
                        this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            this.slideRendererService = (XSlideRenderer)
                UnoRuntime.queryInterface(
                XSlideRenderer.class, service);
        //}
        return this.slideRendererService;
    }
    
    /**
     * Returns the desktop service.
     *
     * @return the desktop service
     */
    private XDesktop getDesktopService() {
        if (this.desktop == null) {
            Object service = null;
            
            try {
                service = this.serviceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            this.desktop = (XDesktop) UnoRuntime.queryInterface(
                XDesktop.class, service);
        }
        return this.desktop;
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
    
    /**
     * An loaded document.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    public class Document implements Comparable<Document> {
        
        /** The presentation document. */
        private XComponent doc = null;
        
        /** The URL of the document. */
        private String url = null;
        
        /** The name of the document. */
        private String name = null;
        
        /**
         * Creates a new instance of an loaded document.
         *
         * @param doc the document compoent
         */
        Document(XComponent doc) {
            XModel xModel = (XModel) UnoRuntime.queryInterface(
                XModel.class, doc);
            int pos = -1;
            OfficeConnection conn = OfficeConnection.this;
            XFileIdentifierConverter nameConverter
                = conn.getFileNameConverterService();
            
            this.doc = doc;
            this.url = xModel.getURL();
            if ("".equals(this.url)) {
                XTitle xTitle = (XTitle) UnoRuntime.queryInterface(
                    XTitle.class, doc);
                
                this.name = xTitle.getTitle();
                return;
            }
            this.name = this.url.substring(this.url.lastIndexOf("/") + 1);
            pos = this.name.lastIndexOf(".");
            if (pos != -1) {
                this.name = this.name.substring(0, pos);
            }
            this.name = nameConverter.getSystemPathFromFileURL(this.name);
            return;
        }
        
        /**
         * Returns a string representation of the document.
         *
         * @return the a string representation of the document, that
         *         is, its file name
         */
        public String toString() {
            return this.name;
        }
        
        /**
         * Returns the document component.
         *
         * @return the document component
         */
        public XComponent getComponent() {
            return this.doc;
        }
        
        /**
         * Indicates whether this document is a presentation document.
         * 
         * @return true if this document is a presentation document; false
         *         otherwise.
         */
        public boolean isPresentation() {
            XServiceInfo xServiceInfo
                = (XServiceInfo) UnoRuntime.queryInterface(
                    XServiceInfo.class, this.doc);
            
            if (xServiceInfo == null) {
                return false;
            }
            if (!xServiceInfo.supportsService(
                    "com.sun.star.presentation.PresentationDocument")) {
                return false;
            }
            return true;
        }
        
        /**
         * Indicates whether some other object is "equal to" this one. 
         *
         * @return true if this object is the same as the obj argument; false
         *         otherwise.
         */
        public boolean equals(Object obj) {
            if (!(obj instanceof Document)) {
                return false;
            }
            if (this.url == null) {
                if (((Document) obj).url == null) {
                    return this.name.equals(((Document) obj).name);
                } else {
                    return false;
                }
            } else {
                return this.url.equals(((Document) obj).url);
            }
        }
        
        /**
         * Compares this object with the specified object for order.  Returns
         * a negative integer, zero, or a positive integer as this object is
         * less than, equal to, or greater than the specified object.
         * <p>
         * The implementor must ensure
         * sgn(x.compareTo(y)) == -sgn(y.compareTo(x)) for all x and y.
         * (This implies that x.compareTo(y) must throw an exception iff
         * y.compareTo(x) throws an exception.)
         * <p>
         * The implementor must also ensure that the relation is transitive:
         * (x.compareTo(y)>0 && y.compareTo(z)>0) implies x.compareTo(z)>0.
         * <p>
         * Finally, the implementor must ensure that x.compareTo(y)==0
         * implies that sgn(x.compareTo(z)) == sgn(y.compareTo(z)), for all
         * z.
         * <p>
         * It is strongly recommended, but not strictly required that
         * (x.compareTo(y)==0) == (x.equals(y)). Generally speaking, any
         * class that implements the Comparable interface and violates this
         * condition should clearly indicate this fact. The recommended
         * language is <q>Note: this class has a natural ordering that is
         * inconsistent with equals.</q>
         * <p>
         * In the foregoing description, the notation sgn(expression)
         * designates the mathematical signum function, which is defined to
         * return one of -1, 0, or 1 according to whether the value of
         * expression is negative, zero or positive.
         *
         * @param o the object to be compared.
         * @return a negative integer, zero, or a positive integer as this
         *         object is less than, equal to, or greater than the
         *         specified object.
         * @throws jav.lang.NullPointerException if the specified object is
         * null
         * @throws java.lang.ClassCastException if the specified object's
         * type prevents it from being compared to this object.
         */
        public int compareTo(Document o) {
            return this.toString().compareTo(o.toString());
        }
    }
}
