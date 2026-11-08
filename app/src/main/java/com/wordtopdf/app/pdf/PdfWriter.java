package com.wordtopdf.app.pdf;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.RectF;
import android.graphics.Typeface;
import android.graphics.pdf.PdfDocument;
import android.os.Build;
import android.text.Layout;
import android.text.SpannableStringBuilder;
import android.text.Spanned;
import android.text.StaticLayout;
import android.text.TextPaint;
import android.text.style.ForegroundColorSpan;
import android.text.style.LeadingMarginSpan;
import android.text.style.RelativeSizeSpan;
import android.text.style.StrikethroughSpan;
import android.text.style.StyleSpan;
import android.text.style.UnderlineSpan;

import com.wordtopdf.app.docx.DocxModel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/** Lays out a {@link DocxModel.Document} onto pages using Android's built-in PdfDocument. */
public final class PdfWriter {

    private final DocxModel.Document doc;
    private final PdfDocument pdf = new PdfDocument();
    private final TextPaint basePaint = new TextPaint(Paint.ANTI_ALIAS_FLAG);
    private final Paint imagePaint = new Paint(Paint.FILTER_BITMAP_FLAG | Paint.ANTI_ALIAS_FLAG);
    private final Paint borderPaint = new Paint(Paint.ANTI_ALIAS_FLAG);

    private final float pageW, pageH, left, top, bottom, contentWidth;

    private PdfDocument.Page page;
    private Canvas canvas;
    private int pageCount;
    private float y;

    /** Writes the PDF and returns the number of pages. */
    public static int write(DocxModel.Document doc, OutputStream out) throws IOException {
        PdfWriter w = new PdfWriter(doc);
        try {
            w.renderAll();
            w.ensurePage();
            w.finishPage();
            w.pdf.writeTo(out);
            out.flush();
            return w.pageCount;
        } finally {
            w.pdf.close();
        }
    }

    private PdfWriter(DocxModel.Document doc) {
        this.doc = doc;
        pageW = doc.pageWidthPt;
        pageH = doc.pageHeightPt;
        float ml = doc.marginLeftPt, mr = doc.marginRightPt, mt = doc.marginTopPt, mb = doc.marginBottomPt;
        if (pageW - ml - mr < 150) { ml = 36; mr = 36; }
        if (pageH - mt - mb < 150) { mt = 36; mb = 36; }
        left = ml;
        top = mt;
        bottom = pageH - mb;
        contentWidth = pageW - ml - mr;

        basePaint.setColor(Color.BLACK);
        basePaint.setTextSize(doc.defaultSizePt);
        basePaint.setTypeface(Typeface.SANS_SERIF);
        borderPaint.setStyle(Paint.Style.STROKE);
        borderPaint.setStrokeWidth(0.6f);
        borderPaint.setColor(Color.DKGRAY);
    }

    // ------------------------------------------------------------------ pages

    private void ensurePage() {
        if (page == null) newPage();
    }

    private void newPage() {
        finishPage();
        pageCount++;
        PdfDocument.PageInfo info =
                new PdfDocument.PageInfo.Builder(Math.round(pageW), Math.round(pageH), pageCount).create();
        page = pdf.startPage(info);
        canvas = page.getCanvas();
        y = top;
    }

    private void finishPage() {
        if (page != null) {
            pdf.finishPage(page);
            page = null;
            canvas = null;
        }
    }

    // ------------------------------------------------------------------ blocks

    private void renderAll() {
        List<DocxModel.Block> blocks = doc.blocks;
        for (int i = 0; i < blocks.size(); i++) {
            DocxModel.Block b = blocks.get(i);
            if (b instanceof DocxModel.Paragraph) {
                drawParagraph((DocxModel.Paragraph) b);
            } else if (b instanceof DocxModel.Image) {
                drawImage((DocxModel.Image) b);
            } else if (b instanceof DocxModel.Table) {
                drawTable((DocxModel.Table) b);
            } else if (b instanceof DocxModel.PageBreak) {
                ensurePage();
                if (i < blocks.size() - 1) newPage();
            }
        }
    }

    private static final class ParaLayout {
        StaticLayout layout;
        float dx;
    }

    private void drawParagraph(DocxModel.Paragraph p) {
        ensurePage();
        ParaLayout pl = layoutParagraph(p, contentWidth);
        if (y > top) y += p.spaceBeforePt;
        drawLayoutAcrossPages(pl.layout, left + pl.dx);
        y += p.spaceAfterPt;
    }

    private ParaLayout layoutParagraph(DocxModel.Paragraph p, float width) {
        float baseSize = p.defaultSizePt > 0 ? p.defaultSizePt : doc.defaultSizePt;
        float indent = Math.max(0, Math.min(p.indentLeftPt, width * 0.5f));
        SpannableStringBuilder sb = new SpannableStringBuilder();
        float hang = 0;
        int firstExtra = 0;

        if (p.listPrefix != null) {
            String prefix = p.listPrefix + "  ";
            sb.append(prefix);
            sb.setSpan(new RelativeSizeSpan(baseSize / doc.defaultSizePt), 0, prefix.length(),
                    Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
            TextPaint tp = new TextPaint(basePaint);
            tp.setTextSize(baseSize);
            hang = tp.measureText(prefix);
            indent = Math.max(indent, hang);
        } else if (p.firstLinePt < 0) {
            hang = Math.min(-p.firstLinePt, indent);
        } else if (p.firstLinePt > 0) {
            firstExtra = (int) Math.min(p.firstLinePt, width * 0.5f);
        }

        for (DocxModel.Run r : p.runs) {
            if (r.text.isEmpty()) continue;
            int start = sb.length();
            sb.append(r.text);
            int end = sb.length();
            float size = r.sizePt > 0 ? r.sizePt : baseSize;
            setSpan(sb, new RelativeSizeSpan(size / doc.defaultSizePt), start, end);
            if (r.bold && r.italic) setSpan(sb, new StyleSpan(Typeface.BOLD_ITALIC), start, end);
            else if (r.bold) setSpan(sb, new StyleSpan(Typeface.BOLD), start, end);
            else if (r.italic) setSpan(sb, new StyleSpan(Typeface.ITALIC), start, end);
            if (r.underline) setSpan(sb, new UnderlineSpan(), start, end);
            if (r.strike) setSpan(sb, new StrikethroughSpan(), start, end);
            if (r.color != null) setSpan(sb, new ForegroundColorSpan(r.color), start, end);
        }

        if (sb.length() == 0) {
            sb.append(" ");
            setSpan(sb, new RelativeSizeSpan(baseSize / doc.defaultSizePt), 0, 1);
        }

        if (hang > 0) {
            sb.setSpan(new LeadingMarginSpan.Standard(0, (int) hang), 0, sb.length(), Spanned.SPAN_INCLUSIVE_INCLUSIVE);
        } else if (firstExtra > 0) {
            sb.setSpan(new LeadingMarginSpan.Standard(firstExtra, 0), 0, sb.length(), Spanned.SPAN_INCLUSIVE_INCLUSIVE);
        }

        float x = indent - hang;
        int w = (int) Math.max(hang + firstExtra + 30, width - x);
        ParaLayout pl = new ParaLayout();
        pl.layout = makeLayout(sb, w, p.align, p.lineSpacing);
        pl.dx = x;
        return pl;
    }

    private static void setSpan(SpannableStringBuilder sb, Object span, int start, int end) {
        sb.setSpan(span, start, end, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
    }

    @SuppressWarnings("deprecation")
    private StaticLayout makeLayout(CharSequence text, int width, int align, float lineMult) {
        Layout.Alignment a;
        switch (align) {
            case DocxModel.ALIGN_CENTER: a = Layout.Alignment.ALIGN_CENTER; break;
            case DocxModel.ALIGN_RIGHT: a = Layout.Alignment.ALIGN_OPPOSITE; break;
            default: a = Layout.Alignment.ALIGN_NORMAL; break;
        }
        if (Build.VERSION.SDK_INT >= 23) {
            StaticLayout.Builder b = StaticLayout.Builder.obtain(text, 0, text.length(), basePaint, width)
                    .setAlignment(a)
                    .setLineSpacing(0f, lineMult)
                    .setIncludePad(true);
            if (align == DocxModel.ALIGN_JUSTIFY && Build.VERSION.SDK_INT >= 26) {
                b.setJustificationMode(Layout.JUSTIFICATION_MODE_INTER_WORD);
            }
            return b.build();
        }
        return new StaticLayout(text, basePaint, width, a, lineMult, 0f, true);
    }

    /** Draws a layout at the current y, continuing onto new pages line by line when needed. */
    private void drawLayoutAcrossPages(StaticLayout layout, float x) {
        int count = layout.getLineCount();
        int line = 0;
        while (line < count) {
            float avail = bottom - y;
            int startTop = layout.getLineTop(line);
            int last = line;
            while (last < count && layout.getLineBottom(last) - startTop <= avail) last++;
            if (last == line) {
                if (y > top) { newPage(); continue; }
                last = line + 1; // line taller than a page: draw it anyway
            }
            int endBottom = layout.getLineBottom(last - 1);
            canvas.save();
            canvas.translate(x, y - startTop);
            canvas.clipRect(-2, startTop, layout.getWidth() + 2, endBottom);
            layout.draw(canvas);
            canvas.restore();
            y += endBottom - startTop;
            line = last;
            if (line < count) newPage();
        }
    }

    // ------------------------------------------------------------------ images

    private void drawImage(DocxModel.Image img) {
        ensurePage();
        BitmapFactory.Options bounds = new BitmapFactory.Options();
        bounds.inJustDecodeBounds = true;
        BitmapFactory.decodeByteArray(img.data, 0, img.data.length, bounds);
        if (bounds.outWidth <= 0 || bounds.outHeight <= 0) return; // e.g. EMF/WMF - not supported

        float w = img.widthPt, h = img.heightPt;
        if (w <= 0 || h <= 0) {
            w = bounds.outWidth * 0.75f;   // assume 96 dpi
            h = bounds.outHeight * 0.75f;
        }
        float scale = Math.min(1f, Math.min(contentWidth / w, (bottom - top) / h));
        w *= scale;
        h *= scale;
        if (y + h > bottom && y > top) newPage();

        int targetW = Math.max(1, (int) (w * 2)), targetH = Math.max(1, (int) (h * 2));
        int sample = 1;
        while (bounds.outWidth / (sample * 2) >= targetW && bounds.outHeight / (sample * 2) >= targetH) sample *= 2;
        BitmapFactory.Options opts = new BitmapFactory.Options();
        opts.inSampleSize = sample;
        Bitmap bmp;
        try {
            bmp = BitmapFactory.decodeByteArray(img.data, 0, img.data.length, opts);
        } catch (OutOfMemoryError oom) {
            return;
        }
        if (bmp == null) return;

        float x = left;
        if (img.align == DocxModel.ALIGN_CENTER) x = left + (contentWidth - w) / 2f;
        else if (img.align == DocxModel.ALIGN_RIGHT) x = left + contentWidth - w;
        canvas.drawBitmap(bmp, null, new RectF(x, y, x + w, y + h), imagePaint);
        bmp.recycle();
        y += h + 4;
    }

    // ------------------------------------------------------------------ tables

    private static final class Cell {
        float x, width;
        final List<ParaLayout> layouts = new ArrayList<>();
        final List<Float> offsets = new ArrayList<>();
        float height;
    }

    private void drawTable(DocxModel.Table t) {
        ensurePage();
        int cols = 0;
        for (List<List<DocxModel.Paragraph>> row : t.rows) cols = Math.max(cols, row.size());
        if (cols == 0) return;

        float[] widths = new float[cols];
        float total = 0;
        boolean useGrid = t.columnWidthsPt.size() == cols;
        for (int c = 0; c < cols; c++) {
            widths[c] = useGrid ? Math.max(t.columnWidthsPt.get(c), 10f) : contentWidth / cols;
            total += widths[c];
        }
        if (total > contentWidth || total <= 0) {
            float f = contentWidth / Math.max(total, 1f);
            for (int c = 0; c < cols; c++) widths[c] *= f;
        }

        final float pad = 4f;
        if (y > top) y += 4;
        for (List<List<DocxModel.Paragraph>> row : t.rows) {
            List<Cell> cells = new ArrayList<>();
            float rowH = 0;
            float cx = left;
            int c = 0;
            while (c < row.size()) {
                int span = 1;
                while (c + span < row.size() && row.get(c + span) == null) span++;
                float cw = 0;
                for (int k = c; k < c + span && k < cols; k++) cw += widths[k];
                Cell cell = new Cell();
                cell.x = cx;
                cell.width = cw;
                List<DocxModel.Paragraph> paras = row.get(c);
                float h = 0;
                if (paras != null) {
                    for (int i = 0; i < paras.size(); i++) {
                        DocxModel.Paragraph p = paras.get(i);
                        if (i > 0) h += Math.min(p.spaceBeforePt, 6f);
                        ParaLayout pl = layoutParagraph(p, Math.max(cw - 2 * pad, 20f));
                        cell.offsets.add(h);
                        cell.layouts.add(pl);
                        h += pl.layout.getHeight();
                        if (i < paras.size() - 1) h += Math.min(p.spaceAfterPt, 6f);
                    }
                }
                cell.height = h + 2 * pad;
                rowH = Math.max(rowH, cell.height);
                cells.add(cell);
                cx += cw;
                c += span;
            }
            rowH = Math.max(rowH, doc.defaultSizePt + 2 * pad);

            if (y + rowH > bottom && y > top) newPage();

            for (Cell cell : cells) {
                for (int i = 0; i < cell.layouts.size(); i++) {
                    ParaLayout pl = cell.layouts.get(i);
                    canvas.save();
                    canvas.clipRect(cell.x, y, cell.x + cell.width, Math.min(y + rowH, bottom));
                    canvas.translate(cell.x + pad + pl.dx, y + pad + cell.offsets.get(i));
                    pl.layout.draw(canvas);
                    canvas.restore();
                }
                canvas.drawRect(cell.x, y, cell.x + cell.width, Math.min(y + rowH, bottom), borderPaint);
            }
            y += rowH;
        }
        y += 6;
    }
}
