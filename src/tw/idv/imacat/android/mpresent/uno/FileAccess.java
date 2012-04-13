/*
 * FileAccess.java
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

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;
import java.util.Random;
import java.util.ResourceBundle;

import com.sun.star.frame.XDesktop;
import com.sun.star.io.XOutputStream;
import com.sun.star.lang.XComponent;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.ucb.XSimpleFileAccess3;
import com.sun.star.uno.UnoRuntime;

/**
 * The file access for the OpenOffice.org remote server.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class FileAccess {
    
    /** The OpenOffice.org remote connection. */
    private OfficeConnection conn = null;
    
    /** The directory listing filter. */
    private Filter dirFilter = null;
    
    /** The file listing filter. */
    private Filter fileFilter = null;
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The file name converter. */
    private XFileIdentifierConverter nameConverter = null;
    
    /** The server file access. */
    private XSimpleFileAccess3 access = null;
    
    /** The parent path segments. */
    private List<PathURL> parents = null;
    
    /**
     * The index of the current working directory in the parent path
     * segments.
     */
    private int pwdParentIndex = -1;
    
    /** The directories. */
    private List<PathURL> dirs = null;
    
    /** The files. */
    private List<PathURL> files = null;
    
    /** The current working directory. */
    private String pwd = null;
    
    /** The temporary working directory. */
    private String tempDir = null;
    
    /** Whether the current working directory is the root directory. */
    private boolean isPwdRoot = false;
    
    /** If the remote host is on MS-Windows. */
    private boolean isMSWindows = false;
    
    /** The characters for the random directory. */
    private static final String randDirChars
        = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    
    /**
     * Creates a new instance of a file access for the OpenOffice.org
     * remote server.
     *
     * @param conn   the OpenOffice.org remote connection
     */
    protected FileAccess(OfficeConnection conn) {
        this.conn = conn;
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.nameConverter = this.conn.getFileNameConverterService();
        this.access = this.conn.getFileAccessService();
        
        if ("/".equals(this.u2p("file:///"))) {
            this.isMSWindows = false;
            this.pwd = this.p2u("/home");
            this.pwd = this.p2u("/home/imacat/Dropbox/samples");
            //this.pwd = this.p2u("/tmp/noaccess/dir");
        } else {
            this.isMSWindows = true;
            //this.pwd = this.p2u("C:\\");
            this.pwd = this.p2u("D:\\imacat\\Dropbox\\samples");
        }
        this.isPwdRoot = this.isRoot(this.pwd);
        if (!this.isPwdRoot) {
            if ("/".equals(this.pwd.substring(this.pwd.length() - 1))) {
                this.pwd = this.pwd.substring(0, this.pwd.length() - 1);
            }
        }
        return;
    }
    
    /**
     * Sets the directory listing filter.
     *
     * @param filter the directory listing filter
     */
    public void setDirFilter(Filter filter) {
        this.dirFilter = filter;
    }
    
    /**
     * Sets the file listing filter.
     *
     * @param filter the file listing filter
     */
    public void setFileFilter(Filter filter) {
        this.fileFilter = filter;
    }
    
    /**
     * Refreshes the remote OpenOffice.org server information.
     *
     * @throws java.io.IOException when error occurs accessing the current
     *         working directory
     */
    public void refresh()
            throws java.io.IOException {
        List<PathURL> temp = null;
        String url = this.pwd;
        String prefix = "";
        String entries[] = new String[0];
        String lastRoot = null;
        String errorMessage = null;
        
        this.parents = new ArrayList<PathURL>();
        this.dirs = new ArrayList<PathURL>();
        this.files = new ArrayList<PathURL>();
        
        // Tries to read the directory, and traces to the parent directory
        // when fails.
        if (this.isMSWindows) {
            lastRoot = this.p2u("C:\\");
        } else {
            lastRoot = this.p2u("/");
        }
        url = this.pwd;
        while (!url.equals(lastRoot)) {
            entries = new String[0];
            try {
                entries = this.access.getFolderContents(url, true);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
                url = this.getOneLevelUpDirectory(url);
                continue;
            } catch (com.sun.star.uno.Exception e) {
                url = this.getOneLevelUpDirectory(url);
                continue;
            }
            if (entries.length == 0) {
                if (this.isMSWindows && this.isRoot(url)) {
                    url = lastRoot;
                }
            }
            break;
        }
        if (url.equals(lastRoot)) {
            try {
                entries = this.access.getFolderContents(url, true);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
            } catch (com.sun.star.uno.Exception e) {
            }
        }
        if (!url.equals(this.pwd)) {
            errorMessage = String.format(_("err_dir"), this.u2p(this.pwd));
            this.pwd = url;
            this.isPwdRoot = this.isRoot(this.pwd);
        }
        
        // Checks the content of the current working directory.
        // Adds the parent directory if not at the root directory.
        if (!this.isPwdRoot) {
            url = url.substring(0, url.lastIndexOf("/"));
            if (this.isRoot(url + "/")) {
                url += "/";
            }
            this.dirs.add(new PathURL(url, "..", FileType.FOLDER));
        }
        for (int i = 0; i < entries.length; i++) {
            String seg = this.u2p(entries[i].substring(
                entries[i].lastIndexOf("/") + 1));
            boolean isEntryFolder = false;
            
            try {
                isEntryFolder = this.access.isFolder(entries[i]);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
                continue;
            } catch (com.sun.star.uno.Exception e) {
                continue;
            }
            if (isEntryFolder) {
                if (this.dirFilter.isAccepted(entries[i])) {
                    this.dirs.add(new PathURL(entries[i], seg,
                        FileType.FOLDER));
                }
            } else {
                if (this.fileFilter.isAccepted(entries[i])) {
                    this.files.add(new PathURL(entries[i], seg));
                }
            }
        }
        Collections.sort(this.dirs);
        Collections.sort(this.files);
        
        // Obtains each path segment of the current working directory.
        url = this.pwd;
        if (!this.isRoot(url)) {
            do {
                int pos = url.lastIndexOf("/");
                String seg = this.u2p(url.substring(pos + 1));
                
                this.parents.add(new PathURL(url, seg));
                url = url.substring(0, pos);
            } while (!this.isRoot(url + "/"));
            url += "/";
        }
        if (this.isMSWindows) {
            int len = url.length();
            
            this.parents.add(new PathURL(url,
                url.substring(len - 3, len - 1)));
        } else {
            this.parents.add(new PathURL(url, "/"));
        }
        
        // Reverses the obtained path segments.
        temp = this.parents;
        this.parents = new ArrayList<PathURL>();
        for (int i = temp.size() - 1; i >= 0; i--) {
            this.parents.add(new PathURL(temp.get(i).getURL(),
                prefix + temp.get(i).toString()));
            prefix += " ";
        }
        this.pwdParentIndex = this.parents.size() - 1;
        
        // Adds the available drives into the path segments.
        if (this.isMSWindows) {
            ArrayList<PathURL> drives = new ArrayList<PathURL>();
            
            for (   int i = "A".codePointAt(0);
                    i <= "Z".codePointAt(0); i++) {
                int codes[] = new int[1];
                String drive = null;
                
                codes[0] = i;
                drive = new String(codes, 0, 1) + ":";
                url = this.p2u(drive + "\\");
                entries = new String[0];
                try {
                    entries = this.access.getFolderContents(url, true);
                } catch (com.sun.star.ucb.CommandAbortedException e) {
                } catch (com.sun.star.uno.Exception e) {
                }
                if (entries.length > 0) {
                    drives.add(new PathURL(url, drive));
                    if (url.equals(this.parents.get(0).getURL())) {
                        for (   int j = 1;
                                j < this.parents.size(); j++) {
                            drives.add(this.parents.get(j));
                        }
                        this.pwdParentIndex = drives.size() - 1;
                    }
                }
            }
            this.parents = drives;
        }
        
        if (errorMessage != null) {
            throw new java.io.IOException(errorMessage);
        }
        return;
    }
    
    /**
     * Checks if an URL represents a folder.
     *
     * @param url URL to be checked
     * @return true, if the given URL represents a folder, otherwise false
     * @throws java.io.IOException when error occurs checking the URL
     */
    public boolean isFolder(String url)
            throws java.io.IOException {
        boolean answer = false;
        
        try {
            answer = this.access.isFolder(url);
        } catch (com.sun.star.ucb.CommandAbortedException e) {
            throw new java.io.IOException(e.getLocalizedMessage());
        } catch (com.sun.star.uno.Exception e) {
            throw new java.io.IOException(e.getLocalizedMessage());
        }
        return answer;
    }
    
    /**
     * Checks if a file is <q>hidden</q>. 
     *
     * @param url URL to be checked
     * @return true, if the given File is <q>hidden</q>, false otherwise 
     */
    public boolean isHidden(String url) {
        if (this.isMSWindows) {
            boolean answer = false;
            
            try {
                answer = this.access.isHidden(url);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
            } catch (com.sun.star.uno.Exception e) {
            }
            return answer;
        }
        
        return url.substring(url.lastIndexOf("/")).startsWith("/.");
    }
    
    /**
     * Changes the current working directory.
     *
     * @param url the URL of the new current working directory
     */
    public void changeDirectory(String url) {
        this.pwd = url;
        this.isPwdRoot = this.isRoot(this.pwd);
        return;
    }
    
    /**
     * Returns the URL of the current working directory.
     *
     * @return the URL of the current working directory
     */
    public String getPwd() {
        return this.pwd;
    }
    
    /**
     * Returns the parent path segments.
     *
     * @return the parent path segments.
     */
    public List<PathURL> getParents() {
        return this.parents;
    }
    
    /**
     * Returns the index of current working directory in the parent path
     * segments.
     *
     * @return the index of current working directory in the parent path
     *         segments
     */
    public int getPwdParentIndex() {
        return this.pwdParentIndex;
    }
    
    /**
     * Returns the directories.
     *
     * @return the directories.
     */
    public List<PathURL> getDirectories() {
        return this.dirs;
    }
    
    /**
     * Returns the files.
     *
     * @return the files.
     */
    public List<PathURL> getFiles() {
        return this.files;
    }
    
    /**
     * Saves the content into a temporary file.
     *
     * @param fileName the file name
     * @param in       the input stream of the file content
     * @return the path URL of the saved file.
     * @throws java.io.IOException if there is an error saving the
     *         content
     */
    public PathURL saveToTemp(String fileName, InputStream in)
            throws java.io.IOException {
        String tempDir = this.getTempDir();
        boolean isFolderExists = false;
        Random random = new Random(Calendar.getInstance().getTimeInMillis());
        char subdir[] = new char[6];
        String url = null;
        XOutputStream out = null;
        int pos = -1;
        
        // Creates a root directory for ourself in the temporary
        // working directory.
        try {
            isFolderExists = this.access.exists(tempDir);
        } catch (com.sun.star.ucb.CommandAbortedException e) {
            throw new java.io.IOException("1: " +
                String.format(_("err_savetemp"), fileName));
        } catch (com.sun.star.uno.Exception e) {
            throw new java.io.IOException("2: " +
                String.format(_("err_savetemp"), fileName));
        }
        if (!isFolderExists) {
            try {
                this.access.createFolder(tempDir);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
                throw new java.io.IOException("3: " + tempDir + ": " +
                    String.format(_("err_savetemp"), fileName));
            } catch (com.sun.star.uno.Exception e) {
                throw new java.io.IOException("4: " + tempDir + ": " +
                    String.format(_("err_savetemp"), fileName));
            }
        }
        
        // Finds an new subdirectory to save the document.
        do {
            for (int i = 0; i < subdir.length; i++) {
                subdir[i] = randDirChars.charAt(
                    random.nextInt(randDirChars.length()));
            }
            url = tempDir + "/" + String.valueOf(subdir);
            try {
                isFolderExists = this.access.exists(url);
            } catch (com.sun.star.ucb.CommandAbortedException e) {
                throw new java.io.IOException("5: " +
                    String.format(_("err_savetemp"), fileName));
            } catch (com.sun.star.uno.Exception e) {
                throw new java.io.IOException("6: " +
                    String.format(_("err_savetemp"), fileName));
            }
        } while (isFolderExists);
        tempDir = url;
        try {
            this.access.createFolder(tempDir);
        } catch (com.sun.star.ucb.CommandAbortedException e) {
            throw new java.io.IOException("7: " + tempDir + ": " +
                String.format(_("err_savetemp"), fileName));
        } catch (com.sun.star.uno.Exception e) {
            throw new java.io.IOException("8: " + tempDir + ": " +
                String.format(_("err_savetemp"), fileName));
        }
        
        // Saves the content to the file.
        url = tempDir + "/" + this.p2u(fileName);
        try {
            out = this.access.openFileWrite(url);
        } catch (com.sun.star.ucb.CommandAbortedException e) {
            throw new java.io.IOException("9: " +
                String.format(_("err_savetemp"), fileName));
        } catch (com.sun.star.uno.Exception e) {
            throw new java.io.IOException("10: " +
                String.format(_("err_savetemp"), fileName));
        }
        while (true) {
            int readSize = -1;
            byte[] data = new byte[65536];
            
            readSize = in.read(data);
            if (readSize <= 0) {
                break;
            }
            try {
                out.writeBytes(Arrays.copyOf(data, readSize));
            } catch (com.sun.star.io.NotConnectedException e) {
                throw new java.io.IOException("11: " +
                    String.format(_("err_savetemp"), fileName));
            } catch (com.sun.star.io.BufferSizeExceededException e) {
                throw new java.io.IOException("12: " +
                    String.format(_("err_savetemp"), fileName));
            } catch (com.sun.star.io.IOException e) {
                throw new java.io.IOException("13: " +
                    String.format(_("err_savetemp"), fileName));
            }
        }
        try {
            out.flush();
            out.closeOutput();
        } catch (com.sun.star.io.NotConnectedException e) {
            throw new java.io.IOException("14: " +
                String.format(_("err_savetemp"), fileName));
        } catch (com.sun.star.io.BufferSizeExceededException e) {
            throw new java.io.IOException("15: " +
                String.format(_("err_savetemp"), fileName));
        } catch (com.sun.star.io.IOException e) {
            throw new java.io.IOException("16: " +
                String.format(_("err_savetemp"), fileName));
        }
        
        pos = fileName.lastIndexOf(".");
        if (pos != -1) {
            fileName = fileName.substring(0, pos);
        }
        return new PathURL(url, fileName);
    }
    
    /**
     * Returns the file URL from the file system path.
     *
     * @param path the file system path
     * @return the file URL
     */
    private String p2u(String path) {
        return this.nameConverter.getFileURLFromSystemPath("", path);
    }
    
    /**
     * Returns the file system path from the file URL.
     *
     * @param url the file URL
     * @return the file system path
     */
    private String u2p(String url) {
        return this.nameConverter.getSystemPathFromFileURL(url);
    }
    
    /**
     * Returns whether the file URL is the root directory.
     *
     * @param url the file URL
     * @return whether the file URL is the root directory
     */
    private boolean isRoot(String url) {
        if (this.isMSWindows) {
            return url.endsWith(":/");
        }
        return "file:///".equals(url);
    }
    
    /**
     * Returns one-level up to the parent directory
     *
     * @param url the original URL
     * @return the one-level up to the parent directory
     */
    private String getOneLevelUpDirectory(String url) {
        String parent = null;
        
        if (this.isMSWindows) {
            if (this.isRoot(url)) {
                return this.p2u("C:\\");
            }
        }
        parent = url.substring(0, url.lastIndexOf("/"));
        if (this.isRoot(parent + "/")) {
            return parent + "/";
        }
        return parent;
    }
    
    /**
     * Returns the URL of the temporary working directory.
     *
     * @return the URL of the temporary working directory
     * @throws java.io.IOException if there is an error finding the temporary
     *         working directory.
     */
    private String getTempDir()
            throws java.io.IOException {
        if (this.tempDir == null) {
            this.tempDir = (String) this.conn.getConfiguration(
                "/org.openoffice.Office.Common/Internal/CurrentTempURL");
            this.tempDir += "/mpresent";
        }
        return this.tempDir;
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
     * A path URL.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    public class PathURL implements Comparable<PathURL> {
        
        /** The actual file URL. */
        private String url = null;
        
        /** The display of the file URL. */
        private String display = null;
        
        /** The file type. */
        private FileType type = null;
        
        /**
         * Creates a new instance of a path URL.
         *
         * @param url    the actual file URL
         * @param display the display of the file URL
         */
        public PathURL(String url, String display) {
            this.url = url;
            this.display = display;
            this.type = FileType.checkType(this.url);
            return;
        }
        
        /**
         * Creates a new instance of a path URL.
         *
         * @param url     the actual file URL
         * @param display the display of the file URL
         * @param type    the file type
         */
        public PathURL(String url, String display, FileType type) {
            this.url = url;
            this.display = display;
            this.type = type;
            return;
        }
        
        /**
         * Returns the actual file URL.
         *
         * @return the actual file URL
         */
        public String getURL() {
            return this.url;
        }
        
        /**
         * Returns a string representation of the file URL.
         *
         * @return the a string representation of the file URL
         */
        public String toString() {
            return this.display;
        }
        
        /**
         * Returns the file type.
         *
         * @return the file type
         */
        public FileType getType() {
            return this.type;
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
        public int compareTo(PathURL o) {
            return this.toString().compareTo(o.toString());
        }
    }
    
    /**
     * A file listing filter.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    public interface Filter {
        
        /**
         * Checks if a file URL is accepted.
         *
         * @param the file URL
         */
        public boolean isAccepted(String url);
    }
    
    /**
     * The file types.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    public enum FileType {
        
        /** A folder. */
        FOLDER ("folder", ":/"),
        
        /** An ODP document. */
        ODP ("odp", ".odp"),
        
        /** An SXI document. */
        SXI ("sxi", ".sxi"),
        
        /** An PPTX document. */
        PPTX ("pptx", ".pptx"),
        
        /** An PPT document. */
        PPT ("ppt", ".ppt");
        
        /** The id of the file type */
        private String id = null;
        
        /** The file suffix. */
        private String suffix = null;
        
        /** The resource ID of the file type description */
        private String resid = null;
        
        /** The localization resources. */
        private ResourceBundle l10n = null;
        
        /**
         * Creates a new instance of a file type.
         *
         * @param id     the ID of the file type
         * @param suffix the file name suffix
         */
        FileType(String id, String suffix) {
            this.id = id;
            this.suffix = suffix;
            this.l10n = ResourceBundle.getBundle(
                this.getClass().getPackage().getName() + ".res.L10n");
            return;
        }
        
        /**
         * Returns the file type.
         *
         * @param url the file URL
         * @return the file type
         */
        public static FileType checkType(String url) {
            String lcUrl = url.toLowerCase();
            
            for (FileType type : FileType.values()) {
                if (lcUrl.endsWith(type.suffix)) {
                    return type;
                }
            }
            return null;
        }
        
        /**
         * Returns the file type description.
         *
         * @return the file type description
         */
        public String getDescription() {
            return _("filedesc_" + this.id);
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
}
