package com.wordtopdf.app;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.ContentResolver;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.provider.DocumentsContract;
import android.provider.OpenableColumns;
import android.view.View;
import android.widget.Button;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import com.wordtopdf.app.docx.DocxModel;
import com.wordtopdf.app.docx.DocxParser;
import com.wordtopdf.app.pdf.PdfWriter;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends Activity {

    private static final int REQ_PICK_WORD = 1;
    private static final int REQ_SAVE_PDF = 2;
    private static final String DOCX_MIME =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final Handler mainHandler = new Handler(Looper.getMainLooper());

    private Uri inputUri;
    private Uri outputUri;
    private String inputName;
    private boolean busy;

    private TextView fileText, statusText;
    private Button pickButton, convertButton, openButton, shareButton;
    private ProgressBar progress;
    private View resultRow;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        fileText = findViewById(R.id.file_text);
        statusText = findViewById(R.id.status_text);
        pickButton = findViewById(R.id.pick_button);
        convertButton = findViewById(R.id.convert_button);
        openButton = findViewById(R.id.open_button);
        shareButton = findViewById(R.id.share_button);
        progress = findViewById(R.id.progress);
        resultRow = findViewById(R.id.result_row);

        pickButton.setOnClickListener(v -> pickWordFile());
        convertButton.setOnClickListener(v -> choosePdfLocation());
        openButton.setOnClickListener(v -> openPdf());
        shareButton.setOnClickListener(v -> sharePdf());

        if (savedInstanceState != null) {
            inputUri = savedInstanceState.getParcelable("input");
            outputUri = savedInstanceState.getParcelable("output");
            inputName = savedInstanceState.getString("name");
        } else {
            handleIncomingIntent(getIntent());
        }
        refreshUi(null);
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        handleIncomingIntent(intent);
        refreshUi(null);
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        out.putParcelable("input", inputUri);
        out.putParcelable("output", outputUri);
        out.putString("name", inputName);
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        executor.shutdown();
    }

    /** Lets the app receive .docx files via "Open with" or "Share" from other apps. */
    @SuppressWarnings("deprecation")
    private void handleIncomingIntent(Intent intent) {
        if (intent == null) return;
        Uri uri = null;
        if (Intent.ACTION_VIEW.equals(intent.getAction())) {
            uri = intent.getData();
        } else if (Intent.ACTION_SEND.equals(intent.getAction())) {
            uri = intent.getParcelableExtra(Intent.EXTRA_STREAM);
        }
        if (uri != null) setInput(uri);
    }

    private void pickWordFile() {
        Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType("*/*");
        i.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{DOCX_MIME, "application/msword", "application/octet-stream"});
        try {
            startActivityForResult(i, REQ_PICK_WORD);
        } catch (ActivityNotFoundException e) {
            toast(getString(R.string.no_file_picker));
        }
    }

    private void choosePdfLocation() {
        if (inputUri == null) return;
        Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType("application/pdf");
        i.putExtra(Intent.EXTRA_TITLE, baseName(inputName) + ".pdf");
        try {
            startActivityForResult(i, REQ_SAVE_PDF);
        } catch (ActivityNotFoundException e) {
            toast(getString(R.string.no_file_picker));
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (resultCode != RESULT_OK || data == null || data.getData() == null) return;
        if (requestCode == REQ_PICK_WORD) {
            setInput(data.getData());
            refreshUi(null);
        } else if (requestCode == REQ_SAVE_PDF) {
            convert(data.getData());
        }
    }

    private void setInput(Uri uri) {
        inputUri = uri;
        outputUri = null;
        inputName = queryName(uri);
    }

    private void convert(Uri target) {
        busy = true;
        outputUri = null;
        refreshUi(getString(R.string.converting));
        final Uri source = inputUri;
        final ContentResolver cr = getContentResolver();
        executor.execute(() -> {
            String error = null;
            int pages = 0;
            try (InputStream in = cr.openInputStream(source)) {
                if (in == null) throw new java.io.IOException("Could not open the Word file.");
                DocxModel.Document doc = DocxParser.parse(in);
                try (OutputStream out = cr.openOutputStream(target, "w")) {
                    if (out == null) throw new java.io.IOException("Could not create the PDF file.");
                    pages = PdfWriter.write(doc, out);
                }
            } catch (OutOfMemoryError oom) {
                error = "The document is too large to convert on this device.";
            } catch (Exception e) {
                error = e.getMessage() != null ? e.getMessage() : e.toString();
            }
            if (error != null) {
                try { DocumentsContract.deleteDocument(cr, target); } catch (Exception ignored) {}
            }
            final String err = error;
            final int pageCount = pages;
            mainHandler.post(() -> {
                busy = false;
                if (err == null) {
                    outputUri = target;
                    refreshUi(getResources().getQuantityString(R.plurals.done, pageCount, pageCount));
                } else {
                    refreshUi(getString(R.string.failed, err));
                }
            });
        });
    }

    private void openPdf() {
        if (outputUri == null) return;
        Intent i = new Intent(Intent.ACTION_VIEW);
        i.setDataAndType(outputUri, "application/pdf");
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        try {
            startActivity(i);
        } catch (ActivityNotFoundException e) {
            toast(getString(R.string.no_pdf_viewer));
        }
    }

    private void sharePdf() {
        if (outputUri == null) return;
        Intent i = new Intent(Intent.ACTION_SEND);
        i.setType("application/pdf");
        i.putExtra(Intent.EXTRA_STREAM, outputUri);
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        startActivity(Intent.createChooser(i, getString(R.string.share_pdf)));
    }

    private void refreshUi(String status) {
        fileText.setText(inputUri == null ? getString(R.string.no_file) : inputName);
        pickButton.setEnabled(!busy);
        convertButton.setEnabled(!busy && inputUri != null);
        progress.setVisibility(busy ? View.VISIBLE : View.GONE);
        resultRow.setVisibility(outputUri != null && !busy ? View.VISIBLE : View.GONE);
        if (status != null) {
            statusText.setText(status);
            statusText.setVisibility(View.VISIBLE);
        } else if (!busy && outputUri == null) {
            statusText.setVisibility(View.GONE);
        }
    }

    private String queryName(Uri uri) {
        String name = null;
        try (Cursor c = getContentResolver().query(uri, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
            if (c != null && c.moveToFirst()) name = c.getString(0);
        } catch (Exception ignored) {}
        if (name == null) name = uri.getLastPathSegment();
        return name != null ? name : "document.docx";
    }

    private static String baseName(String name) {
        if (name == null) return "document";
        int slash = name.lastIndexOf('/');
        if (slash >= 0) name = name.substring(slash + 1);
        int dot = name.lastIndexOf('.');
        return dot > 0 ? name.substring(0, dot) : name;
    }

    private void toast(String msg) {
        Toast.makeText(this, msg, Toast.LENGTH_LONG).show();
    }
}
