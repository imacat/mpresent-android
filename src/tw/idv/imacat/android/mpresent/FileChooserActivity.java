/*
 * FileChooserActivity.java
 *
 * Created on 2012-01-19
 * 
 * Copyright (c) 2011-2012 imacat
 */

/*
 * Copyright (c) 2011-2012 imacat
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
package tw.idv.imacat.android.mpresent;

import java.util.List;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.os.Bundle;
import android.os.StrictMode;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.Button;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ImageView;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Spinner;

import tw.idv.imacat.android.mpresent.uno.FileAccess;
import tw.idv.imacat.android.mpresent.uno.ImpressController;
import tw.idv.imacat.android.mpresent.uno.OfficeConnection;

/**
 * The file chooser for the remote OpenOffice.org server.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class FileChooserActivity extends Activity {
    
    /** The OpenOffice.org Impress Controller. */
    private OfficeConnection conn = null;
    
    /** The file access for the OpenOffice.org remote server. */
    private FileAccess access = null;
    
    /**
     * Called when the activity is starting.  This is where most
     * initialization should go: calling {@link #setContentView(int)} to
     * inflate the activity's UI, using {@link #findViewById(int)} to
     * programmatically interact with widgets in the UI, calling
     * {@link #managedQuery(android.net.Uri,String[],String,String[],String)}
     * to retrieve cursors for data being displayed, etc. 
     *
     * @param savedInstanceState If the activity is being re-initialized
     * after previously being shut down then this Bundle contains the data
     * it most recently supplied in {@link #onSaveInstanceState(Bundle)}.
     * Note: Otherwise it is null.
     */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        TextView txtMessage = null;
        Spinner spnParents = null;
        ListView lstFiles = null;
        
        super.onCreate(savedInstanceState);
        
        // Avoids android.os.NetworkOnMainThreadException.
        // See http://developer.android.com/reference/android/os/NetworkOnMainThreadException.html
        // or .detectAll() for all detectable problems
        StrictMode.setThreadPolicy(new StrictMode.ThreadPolicy.Builder()
            .detectDiskReads()
            .detectDiskWrites()
            .detectNetwork()
            .penaltyLog()
            .build());
        StrictMode.setVmPolicy(new StrictMode.VmPolicy.Builder()
            .detectLeakedSqlLiteObjects()
            .penaltyLog()
            .penaltyDeath()
            .build());
        
        this.setContentView(R.layout.file_chooser);
        
        txtMessage = (TextView) this.findViewById(R.id.message);
        this.conn = RemoteControllerActivity.getSavedConnection();
        this.access = this.conn.getFileAccess();
        this.access.setDirFilter(new DirectoryFilter(this.access));
        this.access.setFileFilter(new PresentationDocumentFilter(
            this.access));
        
        spnParents = (Spinner) this.findViewById(R.id.parents);
        spnParents.setOnItemSelectedListener(new ParentSelected());
        
        lstFiles = (ListView) this.findViewById(R.id.files);
        lstFiles.setOnItemClickListener(new FileClicked());
        
        this.redraw();
        return;
    }
    
    /**
     * Reloads the server status and refreshes the screen.
     *
     * @param btnReload the reload button
     */
    public void reload(View btnReload) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        txtMessage.setText("");
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Redraws the activity.
     *
     */
    public void redraw() {
        List<FileAccess.PathURL> parents = null;
        ArrayAdapter<FileAccess.PathURL> parentsAdapter = null;
        Spinner spnParents = (Spinner) this.findViewById(R.id.parents);
        ArrayAdapter<FileAccess.PathURL> filesAdapter = null;
        ListView lsvFiles = (ListView) this.findViewById(R.id.files);
        
        try {
            this.access.refresh();
        } catch (java.io.IOException e) {
            TextView txtMessage = (TextView) this.findViewById(R.id.message);
            txtMessage.setText(e.getMessage());
        }
        
        // Updates the parent directories list.
        parents = this.access.getParents();
        parentsAdapter = new ArrayAdapter<FileAccess.PathURL>(this,
            android.R.layout.simple_spinner_dropdown_item);
        parentsAdapter.addAll(parents);
        spnParents.setAdapter(parentsAdapter);
        spnParents.setSelection(this.access.getPwdParentIndex());
        
        // Updates the files list.
        filesAdapter = new FilesAdapter(this,
            android.R.layout.simple_list_item_1);
        filesAdapter.addAll(this.access.getDirectories());
        filesAdapter.addAll(this.access.getFiles());
        lsvFiles.setAdapter(filesAdapter);
        return;
    }
    
    /**
     * Enables or disables all the buttons.
     *
     * @param enable whether to enable all the buttons
     */
    private void enableButtons(boolean enable) {
        Button btnReload = (Button) this.findViewById(R.id.reload);
        Spinner spnParents = (Spinner) this.findViewById(R.id.parents);
        ListView lsvFiles = (ListView) this.findViewById(R.id.files);
        
        btnReload.setClickable(enable);
        spnParents.setClickable(enable);
        lsvFiles.setClickable(enable);
        return;
    }
    
    /**
     * The parent selection listener.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class ParentSelected implements AdapterView.OnItemSelectedListener {
        
        /** The activity. */
        private FileChooserActivity act = null;
        
        /**
         * Creates a new instance of a list item click listener.
         *
         */
        public ParentSelected() {
            this.act = FileChooserActivity.this;
        }
        
        /**
         * Callback method to be invoked when an item in this AdapterView has 
         * been clicked.
         * <p>
         * Implementers can call {@link #getItemAtPosition(position)} if they
         * need to access the data associated with the selected item.
         *
         * @param parent   The AdapterView where the selection happened.
         * @param view     The view within the AdapterView that was clicked.
         * @param position The position of the view in the adapter.
         * @param id       The row id of the item that was selected.
         */
        public void onItemSelected(AdapterView<?> parent,
                View view, int position, long id) {
            FileAccess.PathURL path
                = (FileAccess.PathURL) ((Spinner) parent).getSelectedItem();
            
            this.act.enableButtons(false);
            if (!path.getURL().equals(this.act.access.getPwd())) {
                this.act.access.changeDirectory(path.getURL());
                this.act.redraw();
            }
            this.act.enableButtons(true);
            return;
        }
        
        public void onNothingSelected (AdapterView<?> parent) {
            return;
        }
    }
    
    /**
     * The files list view item click listener.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class FileClicked implements AdapterView.OnItemClickListener {
        
        /** The activity. */
        private FileChooserActivity act = null;
        
        /**
         * Creates a new instance of a list item click listener.
         *
         */
        public FileClicked() {
            this.act = FileChooserActivity.this;
        }
        
        /**
         * Callback method to be invoked when an item in this AdapterView has 
         * been clicked.
         * <p>
         * Implementers can call {@link #getItemAtPosition(position)} if they
         * need to access the data associated with the selected item.
         *
         * @param parent   The AdapterView where the click happened.
         * @param view     The view within the AdapterView that was clicked
         *                 (this will be a view provided by the adapter).
         * @param position The position of the view in the adapter.
         * @param id       The row id of the item that was clicked.
         */
        public void onItemClick(AdapterView<?> parent,
                View view, int position, long id) {
            TextView txtMessage
                = (TextView) this.act.findViewById(R.id.message);
            FileAccess.PathURL path
                = (FileAccess.PathURL) parent.getItemAtPosition(position);
            String url = path.getURL();
            boolean isItemFolder = false;
            ImpressController controller = null;
            OfficeConnection.Document doc = null;
            
            this.act.enableButtons(false);
            try {
                isItemFolder = this.act.access.isFolder(url);
            } catch (java.io.IOException e) {
                txtMessage.setText(e.getMessage());
                this.act.redraw();
                this.act.enableButtons(true);
                return;
            }
            if (isItemFolder) {
                txtMessage.setText("");
                this.act.access.changeDirectory(url);
                this.act.redraw();
                this.act.enableButtons(true);
                return;
            }
            
            try {
                doc = this.act.conn.open(url);
            } catch (java.io.IOException e) {
                txtMessage.setText(e.getMessage());
                this.act.redraw();
                this.act.enableButtons(true);
                return;
            } catch (java.lang.IllegalArgumentException e) {
                txtMessage.setText(e.getMessage());
                this.act.redraw();
                this.act.enableButtons(true);
                return;
            }
            if (!doc.isPresentation()) {
                txtMessage.setText(String.format(
                    this.act.getResources().getString(
                        R.string.err_not_presentation), path.toString()));
                this.act.redraw();
                this.act.enableButtons(true);
                return;
            }
            controller = this.act.conn.getImpressController();
            try {
                controller.start(doc);
                txtMessage.setText("");
            } catch (com.sun.star.comp.helper.BootstrapException e) {
                txtMessage.setText(e.getMessage());
            } catch (com.sun.star.connection.NoConnectException e) {
                txtMessage.setText(e.getMessage());
            } catch (com.sun.star.connection.ConnectionSetupException e) {
                txtMessage.setText(e.getMessage());
            } catch (tw.idv.imacat.android.mpresent.uno.ClosedException e) {
                txtMessage.setText(e.getMessage());
            }
            this.act.setResult(Activity.RESULT_OK);
            this.act.finish();
            return;
        }
    }
    
    /**
     * A directory listig filter.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class DirectoryFilter implements FileAccess.Filter {
        
        /** The file access. */
        private FileAccess access = null;
        
        /**
         * Creates a new instance of a directorys listing filter.
         *
         * @param access the file acess.
         */
        public DirectoryFilter(FileAccess access) {
            this.access = access;
        }
        
        /**
         * Checks if a file URL is accepted.
         *
         * @param the file URL
         */
        public boolean isAccepted(String url) {
            return !this.access.isHidden(url);
        }
    }
    
    /**
     * A presentation document file listig filter.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class PresentationDocumentFilter implements FileAccess.Filter {
        
        /** The file access. */
        private FileAccess access = null;
        
        /**
         * Creates a new instance of a presentation document file listing
         * filter.
         *
         * @param access the file acess.
         */
        public PresentationDocumentFilter(FileAccess access) {
            this.access = access;
        }
        
        /**
         * Checks if a file URL is accepted.
         *
         * @param the file URL
         */
        public boolean isAccepted(String url) {
            String lcUrl = null;
            
            if (this.access.isHidden(url)) {
                return false;
            }
            
            lcUrl = url.toLowerCase();
            if (lcUrl.endsWith(".odp") || lcUrl.endsWith(".pptx")
                    || lcUrl.endsWith(".ppt") || lcUrl.endsWith("sxi")) {
                return true;
            }
            return false;
        }
    }
    
    /**
     * A files list adapter.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class FilesAdapter extends ArrayAdapter<FileAccess.PathURL> {
        
        /** The layout inflater. */
        private LayoutInflater inflater = null;
        
        /**
         * Creates a new instance of a files list adapter.
         *
         * @param context            The current context.
         * @param textViewResourceId The resource ID for a layout file
         *                           containing a TextView to use when
         *                           instantiating views. 
         */
        public FilesAdapter(Context context, int textViewResourceId) {
            super(context, textViewResourceId);
            this.inflater = (LayoutInflater) context.getSystemService(
                Context.LAYOUT_INFLATER_SERVICE);
        }
        
        /**
         * Get a View that displays the data at the specified position in the
         * data set.  You can either create a View manually or inflate it
         * from an XML layout file.  When the View is inflated, the parent
         * View (GridView, ListView...) will apply default layout parameters
         * unless you use
         * {@link #inflate(int, android.view.ViewGroup, boolean)} to specify a
         * root view and to prevent attachment to the root.
         *
         * @param position    The position of the item within the adapter's
         *                    data set of the item whose view we want.
         * @param convertView The old view to reuse, if possible.  Note: You
         *                    should check that this view is non-null and of
         *                    an appropriate type before using. If it is not
         *                    possible to convert this view to display the
         *                    correct data, this method can create a new
         *                    view.  Heterogeneous lists can specify their
         *                    number of view types, so that this View is
         *                    always of the right type (see
         *                    {@link #getViewTypeCount()} and
         *                    {@link #getItemViewType(int)}).
         * @param parent      The parent that this view will eventually be
         *                    attached to
         */
        @Override
        public View getView(int position, View convertView,
                ViewGroup parent) {
            View row = this.inflater.inflate(
                R.layout.file_row, parent, false);
            ImageView imgIcon = (ImageView) row.findViewById(R.id.icon);
            TextView txtFile = (TextView) row.findViewById(R.id.file);
            TextView txtDescription = (TextView) row.findViewById(R.id.desc);
            FileAccess.PathURL path = (FileAccess.PathURL)
                ((AdapterView<?>) parent).getItemAtPosition(position);
            FileAccess.FileType type = path.getType();
            
            txtFile.setText(path.toString());
            switch (type) {
            case FOLDER:
                imgIcon.setImageResource(R.drawable.folder);
                break;
            case ODP:
                imgIcon.setImageResource(R.drawable.odp);
                break;
            case PPTX:
                imgIcon.setImageResource(R.drawable.pptx);
                break;
            case PPT:
                imgIcon.setImageResource(R.drawable.ppt);
                break;
            case SXI:
                imgIcon.setImageResource(R.drawable.sxi);
                break;
            }
            txtDescription.setText(type.getDescription());
            return row;
        }
    }
}
