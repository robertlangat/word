package com.wordtopdf.app.docx;

import java.util.ArrayList;
import java.util.List;

/** Simple in-memory representation of a Word document. Units are PDF points (1/72 inch). */
public final class DocxModel {

    private DocxModel() {}

    public interface Block {}

    public static final class Run {
        public String text = "";
        public boolean bold, italic, underline, strike;
        public float sizePt;          // 0 = use document default
        public Integer color;         // ARGB, null = default (black)
    }

    public static final int ALIGN_LEFT = 0, ALIGN_CENTER = 1, ALIGN_RIGHT = 2, ALIGN_JUSTIFY = 3;

    public static final class Paragraph implements Block {
        public final List<Run> runs = new ArrayList<>();
        public int align = ALIGN_LEFT;
        public float indentLeftPt, firstLinePt, spaceBeforePt, spaceAfterPt;
        public float lineSpacing = 1f;  // multiplier
        public String listPrefix;       // e.g. "•" or "1." ; null = not a list item
        public float defaultSizePt;     // size of empty paragraphs / list prefix

        public boolean isEmpty() {
            for (Run r : runs) if (!r.text.isEmpty()) return false;
            return true;
        }

        public String plainText() {
            StringBuilder sb = new StringBuilder();
            for (Run r : runs) sb.append(r.text);
            return sb.toString();
        }
    }

    public static final class Image implements Block {
        public byte[] data;
        public float widthPt, heightPt;
        public int align = ALIGN_LEFT;
    }

    public static final class Table implements Block {
        public final List<Float> columnWidthsPt = new ArrayList<>();
        /** rows -> cells -> paragraphs */
        public final List<List<List<Paragraph>>> rows = new ArrayList<>();
    }

    public static final class PageBreak implements Block {}

    public static final class Document {
        public final List<Block> blocks = new ArrayList<>();
        public float pageWidthPt = 595f;   // A4
        public float pageHeightPt = 842f;
        public float marginTopPt = 72f, marginBottomPt = 72f, marginLeftPt = 72f, marginRightPt = 72f;
        public float defaultSizePt = 11f;
    }
}
