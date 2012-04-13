/*
 * RemoteControllerActivity.java
 *
 * Created on 2011-12-30
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

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Hashtable;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.res.Configuration;
import android.graphics.Bitmap;
import android.graphics.drawable.BitmapDrawable;
import android.net.Uri;
import android.os.Binder;
import android.os.Bundle;
import android.os.StrictMode;
import android.util.TypedValue;
import android.view.Gravity;
import android.view.inputmethod.InputMethodManager;
import android.view.View;
import android.view.WindowManager;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageButton;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Spinner;

import tw.idv.imacat.android.mpresent.uno.FileAccess;
import tw.idv.imacat.android.mpresent.uno.ImpressController;
import tw.idv.imacat.android.mpresent.uno.OfficeConnection;

import java.util.List;
import android.os.IBinder;
import android.os.Parcel;

/**
 * The OpenOffice.org Presentation Remote Controller Activity for Android.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class RemoteControllerActivity extends Activity {
    
    /** The request to open a file. */
    private static final int REQUEST_OPEN_FILE = 0;
    
    /** The internal hosts list.  Should be replaced by caching. */
    private static final Hashtable<String,String> candHosts
    	= new Hashtable<String,String>();
    static {
    };
    
    /** The default remote port.  Should be replaced by caching. */
    private static final int defaultPort = 2002;
    
    /** The OpenOffice.org remote connection. */
    private OfficeConnection conn = null;
    
    /** The OpenOffice.org remote connections. */
    private Hashtable<String,OfficeConnection> conns
        = new Hashtable<String,OfficeConnection>();
    
    /** The Impress Controller. */
    private ImpressController controller = null;
    
    /** The local URI of the presentation document to run. */
    private Uri toRun = null;
    
    /** Whether the local presentation document was run. */
    private boolean isIntentLoaded = false;
    
    /** The saved data to redraw the activity. */
    private static RedrawData savedRedrawData = null;
    
    /** The saved OpenOffice.org remote connection. */
    private static OfficeConnection savedConn = null;
    
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
        ArrayAdapter<String> adapter = null;
        AutoCompleteTextView hostView = null;
        Intent intent = this.getIntent();
        
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
        
        this.setContentView(R.layout.main);
        
        // Saves the intent.
        this.toRun = null;
        if (    !this.isIntentLoaded
                && intent.getAction().equals(Intent.ACTION_VIEW)) {
            this.toRun = intent.getData();
            this.isIntentLoaded = true;
        }
        
        // Sets the auto-complete list.
        adapter = new ArrayAdapter<String>(this,
            android.R.layout.simple_dropdown_item_1line);
        adapter.addAll(this.candHosts.keySet());
        hostView = (AutoCompleteTextView) this.findViewById(R.id.host);
        hostView.setAdapter(adapter);
        // Disables the screen saver.
        this.getWindow().addFlags(
            WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
        
        if (RemoteControllerActivity.savedRedrawData != null) {
            AutoCompleteTextView edtHost = (AutoCompleteTextView)
                this.findViewById(R.id.host);
            EditText edtPort = (EditText) this.findViewById(R.id.port);
            RedrawData redrawData = RemoteControllerActivity.savedRedrawData;
            
            this.enableButtons(false);
            this.conn = redrawData.conn;
            this.conns = redrawData.conns;
            this.controller = redrawData.controller;
            if (this.toRun == null && redrawData.toRun != null) {
                this.toRun = redrawData.toRun;
            }
            edtHost.setText(redrawData.host);
            edtPort.setText(redrawData.port);
            
            // Opens the presentation awaits.
            if (this.toRun != null && this.conn != null) {
                this.runIntented();
                this.toRun = null;
                this.controller.refresh();
                redrawData.presentations
                    = this.controller.getLoadedPresentations();
            }
            
            this.redraw(redrawData);
            RemoteControllerActivity.savedRedrawData = null;
            this.enableButtons(true);
        } else {
            this.enableButtons(false);
            this.redraw();
            this.enableButtons(true);
        }
        return;
    }
    
    /**
     * Called to retrieve per-instance state from an activity before being
     * killed so that the state can be restored in {@link #onCreate(Bundle)}
     * or {@link #onRestoreInstanceState(Bundle)} (the Bundle populated by
     * this method will be passed to both).
     * <p>
     * This method is called before an activity may be killed so that when it
     * comes back some time in the future it can restore its state.  For
     * example, if activity B is launched in front of activity A, and at some
     * point activity A is killed to reclaim resources, activity A will have
     * a chance to save the current state of its user interface via this
     * method so that when the user returns to activity A, the state of the
     * user interface can be restored via {@link #onCreate(Bundle)} or
     * {@link #onRestoreInstanceState(Bundle)}.
     * <p>
     * Do not confuse this method with activity lifecycle callbacks such as
     * {@link #onPause()}, which is always called when an activity is being
     * placed in the background or on its way to destruction, or
     * {@link #onStop()} which is called before destruction.  One example of
     * when {@link #onPause()} and {@link #onStop()} is called and not this
     * method is when a user navigates back from activity B to activity A:
     * there is no need to call {@link #onSaveInstanceState(Bundle)} on B
     * because that particular instance will never be restored, so the system
     * avoids calling it.  An example when {@link #onPause()} is called and
     * not {@link #onSaveInstanceState(Bundle)} is when activity B is
     * launched in front of activity A: the system may avoid calling
     * {@link #onSaveInstanceState(Bundle)} on activity A if it isn't killed
     * during the lifetime of B since the state of the user interface of A
     * will stay intact.
     * <p>
     * The default implementation takes care of most of the UI per-instance
     * state for you by calling {@link #onSaveInstanceState()} on each view
     * in the hierarchy that has an id, and by saving the id of the currently
     * focused view (all of which is restored by the default implementation
     * of {@link #onRestoreInstanceState(Bundle)}).  If you override this
     * method to save additional information not captured by each individual
     * view, you will likely want to call through to the default
     * implementation, otherwise be prepared to save all of the state of each
     * view yourself.
     * <p>
     * If called, this method will occur before {@link #onStop()}.  There are
     * no guarantees about whether it will occur before or after
     * {@link #onPause()}.
     *
     * @param outState Bundle in which to place your saved state.
     */
    @Override
    protected void onSaveInstanceState(Bundle outState) {
        AutoCompleteTextView edtHost = (AutoCompleteTextView)
            this.findViewById(R.id.host);
        EditText edtPort = (EditText) this.findViewById(R.id.port);
        
        super.onSaveInstanceState(outState);
        RemoteControllerActivity.savedRedrawData = this.getRedrawData();
        RemoteControllerActivity.savedRedrawData.host
            = edtHost.getText().toString();
        RemoteControllerActivity.savedRedrawData.port
            = edtPort.getText().toString();
        RemoteControllerActivity.savedRedrawData.toRun = this.toRun;
        return;
    }
    
    /**
     * Called to retrieve per-instance state from an activity before being
     * killed so that the state can be restored in {@link #onCreate(Bundle)}
     * or {@link #onRestoreInstanceState(Bundle)} (the Bundle populated by
     * this method will be passed to both).
     * <p>
     * This method is called between {@link #onStart()} and
     * {@link #onPostCreate(Bundle)}.
     *
     * @param savedInstanceState Bundle in which to place your saved state.
     */
    @Override
    protected void onRestoreInstanceState(Bundle savedInstanceState) {
        super.onRestoreInstanceState(savedInstanceState);
        if (RemoteControllerActivity.savedRedrawData != null) {
            AutoCompleteTextView edtHost = (AutoCompleteTextView)
                this.findViewById(R.id.host);
            EditText edtPort = (EditText) this.findViewById(R.id.port);
            RedrawData redrawData = RemoteControllerActivity.savedRedrawData;
            
            this.enableButtons(false);
            this.conn = redrawData.conn;
            this.conns = redrawData.conns;
            this.controller = redrawData.controller;
            this.toRun = redrawData.toRun;
            edtHost.setText(redrawData.host);
            edtPort.setText(redrawData.port);
            
            // Opens the presentation awaits.
            if (this.toRun != null && this.conn != null) {
                this.runIntented();
                this.toRun = null;
                this.controller.refresh();
                redrawData.presentations
                    = this.controller.getLoadedPresentations();
            }
            
            this.redraw(redrawData);
            RemoteControllerActivity.savedRedrawData = null;
            this.enableButtons(true);
        }
        return;
    }
    
    /**
     * Connects to the OpenOffice.org process.
     *
     * @param btnConnect the connect button
     */
    public void connect(View btnConnect) {
        AutoCompleteTextView edtHost = (AutoCompleteTextView)
            this.findViewById(R.id.host);
        EditText edtPort = (EditText) this.findViewById(R.id.port);
        String host = null;
        int port = -1;
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        
        host = edtHost.getText().toString();
        if (this.candHosts.containsKey(host)) {
            host = this.candHosts.get(host);
        }
        if ("".equals(edtPort.getText().toString())) {
            port = this.defaultPort;
        } else {
            port = Integer.parseInt(edtPort.getText().toString());
        }
        
        this.conn = null;
        if (this.conns.containsKey(host + ":" + port)) {
            this.conn = this.conns.get(host + ":" + port);
            if (this.conn.isConnected()) {
                this.controller = this.conn.getImpressController();
                txtMessage.setText("");
            } else {
                this.conn = null;
            }
        }
        
        if (this.conn == null) {
            this.controller = null;
            if ("".equals(host)) {
                txtMessage.setText(R.string.err_host_empty);
            } else {
                try {
                    this.conn = new OfficeConnection(host, port);
                    this.conns.put(host + ":" + port, this.conn);
                    this.controller = this.conn.getImpressController();
                    txtMessage.setText("");
                } catch (com.sun.star.comp.helper.BootstrapException e) {
                    txtMessage.setText(e.getMessage());
                } catch (com.sun.star.connection.NoConnectException e) {
                    txtMessage.setText(e.getMessage());
                } catch (com.sun.star.connection.ConnectionSetupException e) {
                    txtMessage.setText(e.getMessage());
                }
            }
        }
        
        // Opens the presentation awaits.
        if (this.toRun != null && this.conn != null) {
            this.runIntented();
            this.toRun = null;
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Opens a presentation document.  Starts the FileChooserActivity.
     *
     * @param btnOpen the open button
     */
    public void open(View btnOpen) {
        RemoteControllerActivity.savedConn = this.conn;
        this.startActivityForResult(
            new Intent(this, FileChooserActivity.class), REQUEST_OPEN_FILE);
        return;
    }
    
    /**
     * Called when an activity you launched exits, giving you the requestCode
     * you started it with, the resultCode it returned, and any additional
     * data from it.  The resultCode will be {@link #RESULT_CANCELED} if the
     * activity explicitly returned that, didn't return any result, or
     * crashed during its operation.
     * <p>
     * You will receive this call immediately before {@link #onResume()} when
     * your activity is re-starting.
     *
     * @param requestCode The integer request code originally supplied to
     *                    {@link startActivityForResult()}, allowing you to
     *                    identify who this result came from.
     * @param resultCode  The integer result code returned by the child
     *                    activity through its setResult().
     * @param data        An Intent, which can return result data to the
     *                    caller (various data can be attached to Intent
     *                    <q>extras</q>).
     */
    protected void onActivityResult(int requestCode, int resultCode,
            Intent data) {
        if (requestCode == REQUEST_OPEN_FILE) {
            if (resultCode == Activity.RESULT_OK) {
                this.redraw();
            }
        }
        return;
    }
    
    /**
     * Reloads the server status and refreshes the screen.
     *
     * @param btnReload the reload button
     */
    public void reload(View btnReload) {
        this.enableButtons(false);
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Starts the presentation.
     *
     * @param btnPlay the play button
     */
    public void start(View btnPlay) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        Spinner spnDocument = (Spinner) this.findViewById(R.id.doc);
        OfficeConnection.Document selected
            = (OfficeConnection.Document) spnDocument.getSelectedItem();
        
        this.enableButtons(false);
        if (selected == null) {
            txtMessage.setText(R.string.err_missing_presentation);
        } else {
            try {
                this.controller.start(selected);
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
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Stops the presentation.
     *
     * @param btnStop the stop button
     */
    public void stop(View btnStop) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        try {
            this.controller.stop();
            txtMessage.setText("");
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.NoConnectException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            txtMessage.setText(e.getMessage());
        } catch (tw.idv.imacat.android.mpresent.uno.NoRunningException e) {
                txtMessage.setText(e.getMessage());
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Goes to the previous-previous slide.
     *
     * @param btnPrevPrev the previous-previous button
     */
    public void gotoPreviousPreviousSlide(View btnPrevPrev) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        try {
            this.controller.gotoPreviousPreviousSlide();
            txtMessage.setText("");
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.NoConnectException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            txtMessage.setText(e.getMessage());
        } catch (tw.idv.imacat.android.mpresent.uno.NoRunningException e) {
                txtMessage.setText(e.getMessage());
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Goes to the previous slide.
     *
     * @param btnPrevious the previous button
     */
    public void gotoPreviousSlide(View btnPrevious) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        try {
            this.controller.gotoPreviousSlide();
            txtMessage.setText("");
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.NoConnectException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            txtMessage.setText(e.getMessage());
        } catch (tw.idv.imacat.android.mpresent.uno.NoRunningException e) {
                txtMessage.setText(e.getMessage());
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Goes to the next slide.
     *
     * @param btnNext the next button
     */
    public void gotoNextSlide(View btnNext) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        try {
            this.controller.gotoNextSlide();
            txtMessage.setText("");
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.NoConnectException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            txtMessage.setText(e.getMessage());
        } catch (tw.idv.imacat.android.mpresent.uno.NoRunningException e) {
                txtMessage.setText(e.getMessage());
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Goes to the next-next slide.
     *
     * @param btnNextNext the next-next button
     */
    public void gotoNextNextSlide(View btnNextNext) {
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        
        this.enableButtons(false);
        try {
            this.controller.gotoNextNextSlide();
            txtMessage.setText("");
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.NoConnectException e) {
            txtMessage.setText(e.getMessage());
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            txtMessage.setText(e.getMessage());
        } catch (tw.idv.imacat.android.mpresent.uno.NoRunningException e) {
                txtMessage.setText(e.getMessage());
        }
        this.redraw();
        this.enableButtons(true);
        return;
    }
    
    /**
     * Returns the data to redraw the activity.
     *
     * @return the data to redraw the activity.
     */
    private RedrawData getRedrawData() {
        RedrawData redrawData = new RedrawData();
        LinearLayout previews = (LinearLayout)
            this.findViewById(R.id.previews);
        LinearLayout preview = (LinearLayout) this.findViewById(R.id.preview);
        byte[] dib = new byte[0];
        ByteArrayInputStream in = null;
        BitmapDrawable previewPic = null;
        float ratio = -1, scale = -1;
        
        // Not connected
        if (this.conn == null || !this.conn.isConnected()) {
            redrawData.isConnected = false;
            redrawData.conn = null;
            redrawData.conns = new Hashtable<String,OfficeConnection>();
            redrawData.controller = null;
            return redrawData;
        }
        
        // Shows the opened presentations when connected.
        redrawData.isConnected = true;
        redrawData.conn = this.conn;
        redrawData.conns = this.conns;
        redrawData.controller = this.controller;
        redrawData.controller.refresh();
        
        redrawData.presentations
            = redrawData.controller.getLoadedPresentations();
        if (!redrawData.controller.isRunning()) {
            redrawData.isRunning = false;
            return redrawData;
        }
        
        // Shows the previous and next buttons when a presentation is running.
        redrawData.isRunning = true;
        redrawData.runningIndex = redrawData.controller.getRunningIndex();
        redrawData.curSlide = redrawData.controller.getCurrentSlideIndex();
        redrawData.slideCount = redrawData.controller.getSlideCount();
        
        // Draw the previews
        if (redrawData.curSlide - 2 >= 0) {
            dib = this.controller.createPreview(redrawData.curSlide - 2,
                previews.getWidth() / 12, previews.getHeight() / 15);
            in = new ByteArrayInputStream(dib);
            previewPic = new BitmapDrawable(in);
            redrawData.previewPrev2 = previewPic.getBitmap();
        } else {
            redrawData.previewPrev2 = null;
        }
        if (redrawData.curSlide - 1 >= 0) {
            dib = this.controller.createPreview(redrawData.curSlide - 1,
                previews.getWidth() / 12, previews.getHeight() / 15);
            in = new ByteArrayInputStream(dib);
            previewPic = new BitmapDrawable(in);
            redrawData.previewPrev1 = previewPic.getBitmap();
        } else {
            redrawData.previewPrev1 = null;
        }
        if (redrawData.curSlide + 1 < redrawData.slideCount) {
            dib = this.controller.createPreview(redrawData.curSlide + 1,
                previews.getWidth() / 12, previews.getHeight() / 15);
            in = new ByteArrayInputStream(dib);
            previewPic = new BitmapDrawable(in);
            redrawData.previewNext1 = previewPic.getBitmap();
        } else {
            redrawData.previewNext1 = null;
        }
        if (redrawData.curSlide + 2 < redrawData.slideCount) {
            dib = this.controller.createPreview(redrawData.curSlide + 2,
                previews.getWidth() / 12, previews.getHeight() / 15);
            in = new ByteArrayInputStream(dib);
            previewPic = new BitmapDrawable(in);
            redrawData.previewNext2 = previewPic.getBitmap();
        } else {
            redrawData.previewNext2 = null;
        }
        
        // Draws the preview of the current page.
        dib = redrawData.controller.createPreview(redrawData.curSlide,
            preview.getWidth() / 3, preview.getHeight() / 3);
        in = new ByteArrayInputStream(dib);
        previewPic = new BitmapDrawable(in);
        redrawData.preview = previewPic.getBitmap();
        
        ratio = (float) redrawData.preview.getHeight()
            / (float) redrawData.preview.getWidth();
        if (    ((float) previews.getHeight() / 15) /
                    ((float) previews.getWidth() / 12)
                > ratio) {
            redrawData.previewWidth = previews.getWidth() / 12;
            redrawData.previewHeight
                = (int) (redrawData.previewWidth * ratio);
        } else {
            redrawData.previewHeight = previews.getHeight() / 15;
            redrawData.previewWidth
                = (int) (redrawData.previewHeight / ratio);
        }
        return redrawData;
    }
    
    /**
     * Redraws the activity.
     *
     */
    private void redraw() {
        this.enableButtons(false);
        this.redraw(this.getRedrawData());
        this.enableButtons(true);
        return;
    }
    
    /**
     * Redraws the activity with the given data.
     *
     * @param redrawData the data to redraw the activity.
     */
    private void redraw(RedrawData redrawData) {
        AutoCompleteTextView edtHost = (AutoCompleteTextView)
            this.findViewById(R.id.host);
        InputMethodManager input = (InputMethodManager)
            this.getSystemService(Context.INPUT_METHOD_SERVICE);
        Spinner spnDocument = (Spinner) this.findViewById(R.id.doc);
        Button btnOpen = (Button) this.findViewById(R.id.open);
        Button btnPlay = (Button) this.findViewById(R.id.play);
        Button btnStop = (Button) this.findViewById(R.id.stop);
        TextView txtPageNo = (TextView) this.findViewById(R.id.pageno);
        Button btnPrev2 = (Button) this.findViewById(R.id.prev2);
        Button btnPrev1 = (Button) this.findViewById(R.id.prev1);
        Button btnNext1 = (Button) this.findViewById(R.id.next1);
        Button btnNext2 = (Button) this.findViewById(R.id.next2);
        Button btnPrevious = (Button) this.findViewById(R.id.previous);
        Button btnNext = (Button) this.findViewById(R.id.next);
        LinearLayout previews = (LinearLayout)
            this.findViewById(R.id.previews);
        LinearLayout preview = (LinearLayout) this.findViewById(R.id.preview);
        ArrayAdapter<OfficeConnection.Document> adapter = null;
        int fontSize = -1, width = -1, height = -1;
        float xScale = -1, yScale = -1, scale = -1;
        BitmapDrawable previewPic = null;
        Bitmap bitmap = null;
        
        // Not connected
        if (!redrawData.isConnected) {
            this.conn = null;
            this.controller = null;
            input.showSoftInput(edtHost, InputMethodManager.SHOW_IMPLICIT);
            spnDocument.setVisibility(View.INVISIBLE);
            btnOpen.setVisibility(View.INVISIBLE);
            btnPlay.setVisibility(View.INVISIBLE);
            btnStop.setVisibility(View.INVISIBLE);
            txtPageNo.setVisibility(View.INVISIBLE);
            btnPrev2.setVisibility(View.INVISIBLE);
            btnPrev1.setVisibility(View.INVISIBLE);
            btnNext1.setVisibility(View.INVISIBLE);
            btnNext2.setVisibility(View.INVISIBLE);
            btnPrevious.setVisibility(View.INVISIBLE);
            btnNext.setVisibility(View.INVISIBLE);
            preview.setBackgroundDrawable(null);
            return;
        }
        
        // Shows the opened presentations when connected.
        input.hideSoftInputFromWindow(edtHost.getWindowToken(),
            InputMethodManager.HIDE_IMPLICIT_ONLY);
        
        adapter = new ArrayAdapter<OfficeConnection.Document>(this,
                android.R.layout.simple_spinner_dropdown_item);
        adapter.addAll(redrawData.presentations);
        spnDocument.setAdapter(adapter);
        spnDocument.setSelection(redrawData.runningIndex);
        spnDocument.setVisibility(View.VISIBLE);
        btnOpen.setVisibility(View.VISIBLE);
        btnPlay.setVisibility(View.VISIBLE);
        if (!redrawData.isRunning) {
            btnStop.setVisibility(View.INVISIBLE);
            txtPageNo.setVisibility(View.INVISIBLE);
            btnPrev2.setVisibility(View.INVISIBLE);
            btnPrev1.setVisibility(View.INVISIBLE);
            btnNext1.setVisibility(View.INVISIBLE);
            btnNext2.setVisibility(View.INVISIBLE);
            btnPrevious.setVisibility(View.INVISIBLE);
            btnNext.setVisibility(View.INVISIBLE);
            preview.setBackgroundDrawable(null);
            return;
        }
        
        // Shows the previous and next buttons when a presentation is running.
        fontSize = btnNext.getWidth() < (btnNext.getHeight() * 2) / 3?
            btnNext.getWidth(): (btnNext.getHeight() * 2) / 3;
        btnPrevious.setTextSize(TypedValue.COMPLEX_UNIT_PX, fontSize);
        btnNext.setTextSize(TypedValue.COMPLEX_UNIT_PX, fontSize);
        
        btnStop.setVisibility(View.VISIBLE);
        txtPageNo.setVisibility(View.VISIBLE);
        txtPageNo.setText((redrawData.curSlide + 1) + "/"
            + redrawData.slideCount);
        if (redrawData.previewPrev2 == null) {
            btnPrev2.setVisibility(View.INVISIBLE);
        } else {
            btnPrev2.setVisibility(View.VISIBLE);
        }
        if (redrawData.previewPrev1 == null) {
            btnPrev1.setVisibility(View.INVISIBLE);
            btnPrevious.setVisibility(View.INVISIBLE);
        } else {
            btnPrev1.setVisibility(View.VISIBLE);
            btnPrevious.setVisibility(View.VISIBLE);
        }
        if (redrawData.previewNext1 == null) {
            btnNext.setVisibility(View.INVISIBLE);
            btnNext1.setVisibility(View.INVISIBLE);
        } else {
            btnNext.setVisibility(View.VISIBLE);
            btnNext1.setVisibility(View.VISIBLE);
        }
        if (redrawData.previewNext2 == null) {
            btnNext2.setVisibility(View.INVISIBLE);
        } else {
            btnNext2.setVisibility(View.VISIBLE);
        }
        
        // Draw the previews
        xScale = (previews.getWidth() / 4)
            / redrawData.previewWidth;
        yScale = (previews.getHeight() / 5)
            / redrawData.previewHeight;
        scale = xScale < yScale? xScale: yScale;
        scale = 3;
        width = (int) (redrawData.previewWidth * scale);
        height = (int) (redrawData.previewHeight * scale);
        if (redrawData.previewPrev2 != null) {
            bitmap = Bitmap.createScaledBitmap(redrawData.previewPrev2,
                width, height, false);
            previewPic = new BitmapDrawable(bitmap);
            previewPic.setGravity(Gravity.CENTER);
            btnPrev2.setBackgroundDrawable(previewPic);
        }
        if (redrawData.previewPrev1 != null) {
            bitmap = Bitmap.createScaledBitmap(redrawData.previewPrev1,
                width, height, false);
            previewPic = new BitmapDrawable(bitmap);
            previewPic.setGravity(Gravity.CENTER);
            btnPrev1.setBackgroundDrawable(previewPic);
        }
        if (redrawData.previewNext1 != null) {
            bitmap = Bitmap.createScaledBitmap(redrawData.previewNext1,
                width, height, false);
            previewPic = new BitmapDrawable(bitmap);
            previewPic.setGravity(Gravity.CENTER);
            btnNext1.setBackgroundDrawable(previewPic);
        }
        if (redrawData.previewNext2 != null) {
            bitmap = Bitmap.createScaledBitmap(redrawData.previewNext2,
                width, height, false);
            previewPic = new BitmapDrawable(bitmap);
            previewPic.setGravity(Gravity.CENTER);
            btnNext2.setBackgroundDrawable(previewPic);
        }
        
        // Draws the preview of the current page.
        xScale = (float) preview.getWidth()
            / (float) redrawData.preview.getWidth();
        yScale = (float) preview.getHeight()
            / (float) redrawData.preview.getHeight();
        scale = xScale < yScale? xScale: yScale;
        scale = 3;
        width = (int) (redrawData.preview.getWidth() * scale);
        height = (int) (redrawData.preview.getHeight() * scale);
        bitmap = Bitmap.createScaledBitmap(redrawData.preview,
            width, height, false);
        previewPic = new BitmapDrawable(bitmap);
        previewPic.setGravity(Gravity.CENTER);
        preview.setBackgroundDrawable(previewPic);
        return;
    }
    
    /**
     * Runs the intented local presentation document.
     *
     * @return the loaded presentation document
     */
    private OfficeConnection.Document runIntented() {
        String scheme = this.toRun.getScheme();
        TextView txtMessage = (TextView) this.findViewById(R.id.message);
        OfficeConnection.Document doc = null;
        
        // A remote presentation document.
        if (    "http".equals(scheme) || "https".equals(scheme)
                || "ftp".equals(scheme)) {
            try {
                doc = this.conn.open(this.toRun.toString());
            } catch (java.io.IOException e) {
                txtMessage.setText(e.getMessage());
                return null;
            } catch (java.lang.IllegalArgumentException e) {
                txtMessage.setText(e.getMessage());
                return null;
            }
            return doc;
        }
        
        // A local presentation document.
        if ("file".equals(scheme)) {
            String pathName = this.toRun.getPath();
            String fileName = pathName.substring(
                pathName.lastIndexOf("/") + 1);
            FileInputStream in = null;
            FileAccess access = this.conn.getFileAccess();
            FileAccess.PathURL path = null;
            
            try {
                in = new FileInputStream(pathName);
            } catch (java.io.FileNotFoundException e) {
                txtMessage.setText(String.format(
                    this.getResources().getString(
                        R.string.err_file_not_found), fileName));
                return null;
            }
            try {
                path = access.saveToTemp(fileName, in);
            } catch (java.io.IOException e) {
                txtMessage.setText(e.getMessage());
                return null;
            }
            try {
                doc = this.conn.open(path.getURL());
            } catch (java.io.IOException e) {
                txtMessage.setText(e.getMessage());
                return null;
            } catch (java.lang.IllegalArgumentException e) {
                txtMessage.setText(e.getMessage());
                return null;
            }
            if (!doc.isPresentation()) {
                txtMessage.setText(String.format(
                    this.getResources().getString(
                        R.string.err_not_presentation), fileName));
                return null;
            }
            txtMessage.setText("");
            
            try {
                this.controller.start(doc);
                txtMessage.setText("");
            } catch (com.sun.star.comp.helper.BootstrapException e) {
                txtMessage.setText(e.getMessage());
                return null;
            } catch (com.sun.star.connection.NoConnectException e) {
                txtMessage.setText(e.getMessage());
                return null;
            } catch (com.sun.star.connection.ConnectionSetupException e) {
                txtMessage.setText(e.getMessage());
                return null;
            } catch (tw.idv.imacat.android.mpresent.uno.ClosedException e) {
                txtMessage.setText(e.getMessage());
                return null;
            }
        }
        return doc;
    }
    
    /**
     * Enables or disables all the buttons.
     *
     * @param enable whether to enable all the buttons
     */
    private void enableButtons(boolean enable) {
        Button btnConnect = (Button) this.findViewById(R.id.connect);
        Button btnReload = (Button) this.findViewById(R.id.reload);
        Button btnOpen = (Button) this.findViewById(R.id.open);
        Button btnPlay = (Button) this.findViewById(R.id.play);
        Button btnStop = (Button) this.findViewById(R.id.stop);
        Button btnPrev2 = (Button) this.findViewById(R.id.prev2);
        Button btnPrev1 = (Button) this.findViewById(R.id.prev1);
        Button btnNext1 = (Button) this.findViewById(R.id.next1);
        Button btnNext2 = (Button) this.findViewById(R.id.next2);
        Button btnPrevious = (Button) this.findViewById(R.id.previous);
        Button btnNext = (Button) this.findViewById(R.id.next);
        
        btnConnect.setClickable(enable);
        btnReload.setClickable(enable);
        btnOpen.setClickable(enable);
        btnPlay.setClickable(enable);
        btnStop.setClickable(enable);
        btnPrev2.setClickable(enable);
        btnPrev1.setClickable(enable);
        btnNext1.setClickable(enable);
        btnNext2.setClickable(enable);
        btnPrevious.setClickable(enable);
        btnNext.setClickable(enable);
        return;
    }
    
    /**
     * Returns the saved OpenOffice.org remote connection.
     *
     * @return the saved OpenOffice.org remote connection
     */
    protected static OfficeConnection getSavedConnection() {
        return RemoteControllerActivity.savedConn;
    }
    
    /**
     * The data to redraw an activity.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class RedrawData {
        
        /** The content of the host field. */
        private String host = null;
        
        /** The content of the port field. */
        private String port = null;
        
        /** The OpenOffice.org remote connection. */
        private OfficeConnection conn = null;
        
        /** The OpenOffice.org remote connections. */
        private Hashtable<String,OfficeConnection> conns = null;
        
        /** The OpenOffice.org Impress controller. */
        private ImpressController controller = null;
        
        /** The local URI of the presentation document to run. */
        private Uri toRun = null;
        
        /** Whether the connection is on and alive. */
        private boolean isConnected = false;
        
        /** Whether a presentation is running. */
        private boolean isRunning = false;
        
        /** The opened presentation documents. */
        private ArrayList<OfficeConnection.Document> presentations = null;
        
        /**
         * The index of the running presentation in the opened presentations.
         */
        private int runningIndex = -1;
        
        /** The index of the current slide of the running presentation. */
        private int curSlide = -1;
        
        /** The number of slides of the running presentation. */
        private int slideCount = -1;
        
        /** The bitmap of the preview of the previous-previous slide. */
        private Bitmap previewPrev2 = null;
        
        /** The bitmap of the preview of the previous slide. */
        private Bitmap previewPrev1 = null;
        
        /** The bitmap of the preview of the next slide. */
        private Bitmap previewNext1 = null;
        
        /** The bitmap of the preview of the next-next slide. */
        private Bitmap previewNext2 = null;
        
        /** The bitmap of the preview of the current slide. */
        private Bitmap preview = null;
        
        /** The width of the preview of the next/previous slide. */
        private int previewWidth = -1;
        
        /** The height of the preview of the next/previous slide. */
        private int previewHeight = -1;
        
        /**
         * Creates a new instance of the data to redraw an activity.
         *
         */
        private RedrawData() {
        }
    }
    
    /**
     * The test data to redraw an activity.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.1.0
     */
    private class TestRedrawData extends Binder {
        
        /** The OpenOffice.org remote connection. */
        private OfficeConnection conn = null;
        
        /**
         * Creates a new instance of the data to redraw an activity.
         *
         */
        private TestRedrawData() {
        }
    }
}
