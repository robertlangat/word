package com.wordtopdf.app.docx;

import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

/**
 * Lightweight .docx reader. Pure Java (java.util.zip + DOM) so it runs on Android
 * without third-party libraries. Supports paragraphs, bold/italic/underline/strike,
 * font sizes, colours, alignment, spacing, indentation, headings, bullet and numbered
 * lists, tables, inline images, line and page breaks, and the document page size.
 */
public final class DocxParser {

    private final Map<String, byte[]> files = new HashMap<>();
    private final Map<String, Props> styles = new HashMap<>();
    private final Map<String, String> styleBasedOn = new HashMap<>();
    private final Map<String, String> styleNames = new HashMap<>();
    private final Map<String, String> relTargets = new HashMap<>();
    private final Map<String, String> numToAbstract = new HashMap<>();
    private final Map<String, Map<Integer, NumLevel>> abstractLevels = new HashMap<>();
    private final Map<String, int[]> numCounters = new HashMap<>();
    private Props docDefaults = new Props();
    private String defaultParaStyle;
    private DocxModel.Document doc;

    public static DocxModel.Document parse(InputStream in) throws IOException {
        return new DocxParser().run(in);
    }

    // ------------------------------------------------------------------ entry

    private DocxModel.Document run(InputStream in) throws IOException {
        readZip(in);
        byte[] main = files.get("word/document.xml");
        if (main == null) {
            throw new IOException("This is not a valid .docx file. Old .doc files are not supported - "
                    + "open it in Word/Google Docs and save as .docx first.");
        }
        doc = new DocxModel.Document();
        if (files.containsKey("word/styles.xml")) parseStyles(xml(files.get("word/styles.xml")));
        if (files.containsKey("word/numbering.xml")) parseNumbering(xml(files.get("word/numbering.xml")));
        if (files.containsKey("word/_rels/document.xml.rels")) parseRels(xml(files.get("word/_rels/document.xml.rels")));
        if (docDefaults.size != null) doc.defaultSizePt = docDefaults.size;

        Element root = xml(main);
        Element body = child(root, "body");
        if (body == null) throw new IOException("Document has no body.");
        parseBlockContainer(body, doc.blocks);
        return doc;
    }

    private void readZip(InputStream in) throws IOException {
        ZipInputStream zis = new ZipInputStream(new BufferedInputStream(in));
        ZipEntry e;
        boolean any = false;
        while ((e = zis.getNextEntry()) != null) {
            any = true;
            String n = e.getName();
            if (n.startsWith("/")) n = n.substring(1);
            if (n.equals("word/document.xml") || n.equals("word/styles.xml")
                    || n.equals("word/numbering.xml") || n.equals("word/_rels/document.xml.rels")
                    || n.startsWith("word/media/")) {
                files.put(n, readAll(zis));
            }
        }
        if (!any) {
            throw new IOException("This is not a valid .docx file. Old .doc files are not supported - "
                    + "open it in Word/Google Docs and save as .docx first.");
        }
    }

    // ------------------------------------------------------------------ body

    private void parseBlockContainer(Element container, List<DocxModel.Block> out) {
        for (Element el : elements(container)) {
            switch (local(el)) {
                case "p": parseParagraph(el, out); break;
                case "tbl": parseTable(el, out); break;
                case "sdt": {
                    Element c = child(el, "sdtContent");
                    if (c != null) parseBlockContainer(c, out);
                    break;
                }
                case "customXml": parseBlockContainer(el, out); break;
                case "sectPr": parseSection(el); break;
                default: break;
            }
        }
    }

    private void parseSection(Element sectPr) {
        Element sz = child(sectPr, "pgSz");
        if (sz != null) {
            float w = twips(attr(sz, "w"), 0), h = twips(attr(sz, "h"), 0);
            if (w > 50 && h > 50) { doc.pageWidthPt = w; doc.pageHeightPt = h; }
        }
        Element m = child(sectPr, "pgMar");
        if (m != null) {
            doc.marginTopPt = clampMargin(Math.abs(twips(attr(m, "top"), 72)));
            doc.marginBottomPt = clampMargin(Math.abs(twips(attr(m, "bottom"), 72)));
            doc.marginLeftPt = clampMargin(twips(attr(m, "left"), 72));
            doc.marginRightPt = clampMargin(twips(attr(m, "right"), 72));
        }
    }

    private static float clampMargin(float v) { return Math.max(18f, Math.min(v, 144f)); }

    /** Mutable state while walking a paragraph (it may be split by images / page breaks). */
    private static final class ParaCtx {
        List<DocxModel.Block> out;
        Props props;
        DocxModel.Paragraph current;
        boolean split;
    }

    private void parseParagraph(Element p, List<DocxModel.Block> out) {
        Element pPr = child(p, "pPr");
        String styleId = pPr != null ? attr(child(pPr, "pStyle"), "val") : null;
        if (styleId == null) styleId = defaultParaStyle;

        Props eff = new Props();
        eff.mergeFrom(docDefaults);
        eff.mergeFrom(resolveStyle(styleId));
        if (pPr != null) {
            Props direct = new Props();
            readPPr(pPr, direct);  // pPr/rPr is paragraph-mark formatting and is intentionally not applied to text
            eff.mergeFrom(direct);
        }
        applyHeadingFallback(styleId, eff);

        ParaCtx ctx = new ParaCtx();
        ctx.out = out;
        ctx.props = eff;
        ctx.current = newParagraph(eff);
        ctx.current.listPrefix = listPrefix(eff);
        if (ctx.current.listPrefix != null) {
            NumLevel lvl = numLevel(eff.numId, eff.ilvl);
            if (lvl != null && lvl.indLeft != null && (pPr == null || !hasDirectIndent(pPr))) {
                ctx.current.indentLeftPt = lvl.indLeft;
            }
            if (ctx.current.indentLeftPt < 18f) ctx.current.indentLeftPt = 18f;
        }

        walkInline(p, ctx, eff);

        if (!(ctx.split && ctx.current.isEmpty())) out.add(ctx.current);
    }

    private static boolean hasDirectIndent(Element pPr) {
        Element ind = child(pPr, "ind");
        return ind != null && (attr(ind, "left") != null || attr(ind, "start") != null);
    }

    private DocxModel.Paragraph newParagraph(Props eff) {
        DocxModel.Paragraph para = new DocxModel.Paragraph();
        para.align = eff.jc != null ? eff.jc : DocxModel.ALIGN_LEFT;
        para.indentLeftPt = eff.indLeft != null ? Math.max(0, eff.indLeft) : 0;
        para.firstLinePt = eff.firstLine != null ? eff.firstLine : 0;
        para.spaceBeforePt = eff.spBefore != null ? eff.spBefore : 0;
        para.spaceAfterPt = eff.spAfter != null ? eff.spAfter : 0;
        float size = eff.size != null ? eff.size : doc.defaultSizePt;
        para.defaultSizePt = size;
        if (eff.lineMult != null) para.lineSpacing = eff.lineMult;
        else if (eff.linePt != null) para.lineSpacing = eff.linePt / (size * 1.17f);
        para.lineSpacing = Math.max(0.7f, Math.min(para.lineSpacing, 3f));
        return para;
    }

    private void walkInline(Element parent, ParaCtx ctx, Props paraProps) {
        for (Element el : elements(parent)) {
            switch (local(el)) {
                case "r": parseRun(el, ctx, paraProps); break;
                case "hyperlink": case "ins": case "smartTag": case "fldSimple":
                case "customXml": case "moveTo": case "bdo": case "dir":
                    walkInline(el, ctx, paraProps); break;
                case "sdt": {
                    Element c = child(el, "sdtContent");
                    if (c != null) walkInline(c, ctx, paraProps);
                    break;
                }
                default: break; // pPr, del, bookmarks, proofErr, ...
            }
        }
    }

    private void parseRun(Element r, ParaCtx ctx, Props paraProps) {
        Props rp = new Props();
        rp.mergeFrom(paraProps);
        Element rPr = child(r, "rPr");
        if (rPr != null) {
            String rStyle = attr(child(rPr, "rStyle"), "val");
            if (rStyle != null) rp.mergeRunFrom(resolveStyle(rStyle));
            Props direct = new Props();
            readRPr(rPr, direct);
            rp.mergeRunFrom(direct);
        }
        StringBuilder text = new StringBuilder();
        processRunContent(r, ctx, rp, text);
        flushText(ctx, rp, text);
    }

    private void processRunContent(Element r, ParaCtx ctx, Props rp, StringBuilder text) {
        for (Element el : elements(r)) {
            switch (local(el)) {
                case "t": text.append(el.getTextContent()); break;
                case "tab": case "ptab": text.append('\t'); break;
                case "cr": text.append('\n'); break;
                case "noBreakHyphen": text.append('-'); break;
                case "sym": text.append('•'); break;
                case "br": {
                    if ("page".equals(attr(el, "type"))) {
                        flushText(ctx, rp, text);
                        splitParagraph(ctx, new DocxModel.PageBreak());
                    } else {
                        text.append('\n');
                    }
                    break;
                }
                case "drawing": case "pict": case "object": {
                    DocxModel.Image img = readImage(el);
                    if (img != null) {
                        flushText(ctx, rp, text);
                        img.align = ctx.current.align;
                        splitParagraph(ctx, img);
                    }
                    break;
                }
                case "AlternateContent": {
                    Element choice = child(el, "Choice");
                    if (choice == null) choice = child(el, "Fallback");
                    if (choice != null) processRunContent(choice, ctx, rp, text);
                    break;
                }
                default: break;
            }
        }
    }

    private void flushText(ParaCtx ctx, Props rp, StringBuilder text) {
        if (text.length() == 0) return;
        DocxModel.Run run = new DocxModel.Run();
        run.text = rp.caps ? text.toString().toUpperCase(Locale.ROOT) : text.toString();
        run.bold = Boolean.TRUE.equals(rp.b);
        run.italic = Boolean.TRUE.equals(rp.i);
        run.underline = Boolean.TRUE.equals(rp.u);
        run.strike = Boolean.TRUE.equals(rp.strike);
        run.sizePt = rp.size != null ? rp.size : 0;
        run.color = rp.color;
        if (!Boolean.TRUE.equals(rp.hidden)) ctx.current.runs.add(run);
        text.setLength(0);
    }

    /** Emits the paragraph so far, then the given block, and starts a continuation paragraph. */
    private void splitParagraph(ParaCtx ctx, DocxModel.Block block) {
        if (!ctx.current.isEmpty()) {
            ctx.current.spaceAfterPt = 0;
            ctx.out.add(ctx.current);
        }
        ctx.out.add(block);
        DocxModel.Paragraph next = newParagraph(ctx.props);
        next.indentLeftPt = ctx.current.indentLeftPt;
        next.spaceBeforePt = 0;
        ctx.current = next;
        ctx.split = true;
    }

    // ------------------------------------------------------------------ images

    private DocxModel.Image readImage(Element container) {
        String rid = null;
        float w = 0, h = 0;
        Element blip = descendant(container, "blip");
        if (blip != null) rid = attr(blip, "embed");
        if (rid == null) {
            Element vml = descendant(container, "imagedata");
            if (vml != null) rid = attr(vml, "id");
        }
        if (rid == null) return null;
        Element extent = descendant(container, "extent");
        if (extent != null) {
            w = emu(attr(extent, "cx"));
            h = emu(attr(extent, "cy"));
        }
        String target = relTargets.get(rid);
        if (target == null) return null;
        byte[] data = files.get(resolvePath(target));
        if (data == null) return null;
        DocxModel.Image img = new DocxModel.Image();
        img.data = data;
        img.widthPt = w;
        img.heightPt = h;
        return img;
    }

    private static String resolvePath(String target) {
        if (target.startsWith("/")) return target.substring(1);
        String path = "word/" + target;
        while (path.contains("/../")) {
            path = path.replaceFirst("[^/]+/\\.\\./", "");
        }
        return path;
    }

    // ------------------------------------------------------------------ tables

    private void parseTable(Element tbl, List<DocxModel.Block> out) {
        DocxModel.Table table = new DocxModel.Table();
        Element grid = child(tbl, "tblGrid");
        if (grid != null) {
            for (Element col : children(grid, "gridCol")) table.columnWidthsPt.add(twips(attr(col, "w"), 0));
        }
        for (Element tr : rowsOf(tbl)) {
            List<List<DocxModel.Paragraph>> row = new ArrayList<>();
            for (Element tc : cellsOf(tr)) {
                List<DocxModel.Block> blocks = new ArrayList<>();
                parseBlockContainer(tc, blocks);
                List<DocxModel.Paragraph> cell = new ArrayList<>();
                flattenToParagraphs(blocks, cell);
                Element tcPr = child(tc, "tcPr");
                int span = 1;
                if (tcPr != null) {
                    String gs = attr(child(tcPr, "gridSpan"), "val");
                    if (gs != null) try { span = Math.max(1, Integer.parseInt(gs)); } catch (NumberFormatException ignored) {}
                }
                row.add(cell);
                // pad spanned cells with empty placeholders so columns stay aligned
                for (int i = 1; i < span; i++) row.add(null);
            }
            if (!row.isEmpty()) table.rows.add(row);
        }
        if (!table.rows.isEmpty()) out.add(table);
    }

    private static List<Element> rowsOf(Element tbl) {
        List<Element> rows = new ArrayList<>();
        for (Element el : elements(tbl)) {
            if (local(el).equals("tr")) rows.add(el);
            else if (local(el).equals("sdt") || local(el).equals("customXml")) {
                Element c = local(el).equals("sdt") ? child(el, "sdtContent") : el;
                if (c != null) rows.addAll(rowsOf(c));
            }
        }
        return rows;
    }

    private static List<Element> cellsOf(Element tr) {
        List<Element> cells = new ArrayList<>();
        for (Element el : elements(tr)) {
            if (local(el).equals("tc")) cells.add(el);
            else if (local(el).equals("sdt") || local(el).equals("customXml")) {
                Element c = local(el).equals("sdt") ? child(el, "sdtContent") : el;
                if (c != null) cells.addAll(cellsOf(c));
            }
        }
        return cells;
    }

    private static void flattenToParagraphs(List<DocxModel.Block> blocks, List<DocxModel.Paragraph> out) {
        for (DocxModel.Block b : blocks) {
            if (b instanceof DocxModel.Paragraph) {
                out.add((DocxModel.Paragraph) b);
            } else if (b instanceof DocxModel.Table) {
                for (List<List<DocxModel.Paragraph>> row : ((DocxModel.Table) b).rows) {
                    for (List<DocxModel.Paragraph> cell : row) if (cell != null) out.addAll(cell);
                }
            }
        }
    }

    // ------------------------------------------------------------------ styles

    private static final class Props {
        Boolean b, i, u, strike, hidden;
        boolean caps;
        Float size;
        Integer color;
        Integer jc;
        Float spBefore, spAfter, lineMult, linePt, indLeft, firstLine;
        String numId;
        Integer ilvl;

        void mergeFrom(Props o) {
            if (o == null) return;
            mergeRunFrom(o);
            if (o.jc != null) jc = o.jc;
            if (o.spBefore != null) spBefore = o.spBefore;
            if (o.spAfter != null) spAfter = o.spAfter;
            if (o.lineMult != null) { lineMult = o.lineMult; linePt = null; }
            if (o.linePt != null) { linePt = o.linePt; lineMult = null; }
            if (o.indLeft != null) indLeft = o.indLeft;
            if (o.firstLine != null) firstLine = o.firstLine;
            if (o.numId != null) numId = o.numId;
            if (o.ilvl != null) ilvl = o.ilvl;
        }

        void mergeRunFrom(Props o) {
            if (o == null) return;
            if (o.b != null) b = o.b;
            if (o.i != null) i = o.i;
            if (o.u != null) u = o.u;
            if (o.strike != null) strike = o.strike;
            if (o.hidden != null) hidden = o.hidden;
            if (o.caps) caps = true;
            if (o.size != null) size = o.size;
            if (o.color != null) color = o.color;
        }
    }

    private void parseStyles(Element root) {
        Element dd = child(root, "docDefaults");
        if (dd != null) {
            Element rpd = child(dd, "rPrDefault");
            if (rpd != null && child(rpd, "rPr") != null) readRPr(child(rpd, "rPr"), docDefaults);
            Element ppd = child(dd, "pPrDefault");
            if (ppd != null && child(ppd, "pPr") != null) readPPr(child(ppd, "pPr"), docDefaults);
        }
        for (Element st : children(root, "style")) {
            String id = attr(st, "styleId");
            if (id == null) continue;
            Props p = new Props();
            Element pPr = child(st, "pPr");
            if (pPr != null) readPPr(pPr, p);
            Element rPr = child(st, "rPr");
            if (rPr != null) readRPr(rPr, p);
            styles.put(id, p);
            String based = attr(child(st, "basedOn"), "val");
            if (based != null) styleBasedOn.put(id, based);
            String name = attr(child(st, "name"), "val");
            if (name != null) styleNames.put(id, name.toLowerCase(Locale.ROOT));
            String def = attr(st, "default");
            if ("paragraph".equals(attr(st, "type")) && ("1".equals(def) || "true".equals(def))) defaultParaStyle = id;
        }
    }

    private Props resolveStyle(String id) {
        if (id == null) return null;
        List<Props> chain = new ArrayList<>();
        String cur = id;
        for (int depth = 0; cur != null && depth < 12; depth++) {
            Props p = styles.get(cur);
            if (p == null) break;
            chain.add(0, p);
            cur = styleBasedOn.get(cur);
        }
        Props result = new Props();
        for (Props p : chain) result.mergeFrom(p);
        return result;
    }

    private static final Pattern HEADING = Pattern.compile("(?i)heading\\s*([1-9])");

    private void applyHeadingFallback(String styleId, Props eff) {
        if (styleId == null) return;
        String name = styleNames.containsKey(styleId) ? styleNames.get(styleId) : styleId.toLowerCase(Locale.ROOT);
        boolean ownSize = false;
        for (String s = styleId; s != null && !s.equals(defaultParaStyle); s = styleBasedOn.get(s)) {
            Props p = styles.get(s);
            if (p != null && p.size != null) { ownSize = true; break; }
        }
        if (ownSize && styles.containsKey(styleId)) return;
        Matcher m = HEADING.matcher(name);
        if (m.find() || HEADING.matcher(styleId).find()) {
            Matcher mm = HEADING.matcher(name);
            int level = mm.find() ? Integer.parseInt(mm.group(1)) : 1;
            float[] sizes = {20, 16, 14, 13, 12, 12, 11, 11, 11};
            eff.size = sizes[Math.min(level, 9) - 1];
            eff.b = true;
            if (eff.spBefore == null || eff.spBefore < 6) eff.spBefore = 12f;
            if (eff.spAfter == null || eff.spAfter < 3) eff.spAfter = 4f;
        } else if (name.equals("title")) {
            eff.size = 26f;
            if (eff.spAfter == null) eff.spAfter = 8f;
        }
    }

    private void readRPr(Element rPr, Props p) {
        for (Element el : elements(rPr)) {
            switch (local(el)) {
                case "b": p.b = toggle(el); break;
                case "i": p.i = toggle(el); break;
                case "strike": case "dstrike": p.strike = toggle(el); break;
                case "vanish": p.hidden = toggle(el); break;
                case "caps": p.caps = toggle(el); break;
                case "u": {
                    String v = attr(el, "val");
                    p.u = v == null || !"none".equals(v);
                    break;
                }
                case "sz": {
                    String v = attr(el, "val");
                    if (v != null) try { p.size = Float.parseFloat(v) / 2f; } catch (NumberFormatException ignored) {}
                    break;
                }
                case "color": {
                    String v = attr(el, "val");
                    if (v != null && v.length() == 6 && !"auto".equalsIgnoreCase(v)) {
                        try { p.color = 0xFF000000 | Integer.parseInt(v, 16); } catch (NumberFormatException ignored) {}
                    }
                    break;
                }
                default: break;
            }
        }
    }

    private void readPPr(Element pPr, Props p) {
        Element jc = child(pPr, "jc");
        if (jc != null) {
            String v = attr(jc, "val");
            if (v != null) {
                switch (v) {
                    case "center": p.jc = DocxModel.ALIGN_CENTER; break;
                    case "right": case "end": p.jc = DocxModel.ALIGN_RIGHT; break;
                    case "both": case "distribute": p.jc = DocxModel.ALIGN_JUSTIFY; break;
                    default: p.jc = DocxModel.ALIGN_LEFT; break;
                }
            }
        }
        Element sp = child(pPr, "spacing");
        if (sp != null) {
            if (attr(sp, "before") != null) p.spBefore = twips(attr(sp, "before"), 0);
            if (attr(sp, "after") != null) p.spAfter = twips(attr(sp, "after"), 0);
            String line = attr(sp, "line");
            if (line != null) {
                try {
                    float v = Float.parseFloat(line);
                    String rule = attr(sp, "lineRule");
                    if (rule == null || "auto".equals(rule)) { p.lineMult = v / 240f; p.linePt = null; }
                    else { p.linePt = v / 20f; p.lineMult = null; }
                } catch (NumberFormatException ignored) {}
            }
        }
        Element ind = child(pPr, "ind");
        if (ind != null) {
            String left = attr(ind, "left");
            if (left == null) left = attr(ind, "start");
            if (left != null) p.indLeft = twips(left, 0);
            if (attr(ind, "firstLine") != null) p.firstLine = twips(attr(ind, "firstLine"), 0);
            if (attr(ind, "hanging") != null) p.firstLine = -twips(attr(ind, "hanging"), 0);
        }
        Element numPr = child(pPr, "numPr");
        if (numPr != null) {
            String numId = attr(child(numPr, "numId"), "val");
            if (numId != null) p.numId = numId;
            String ilvl = attr(child(numPr, "ilvl"), "val");
            if (ilvl != null) try { p.ilvl = Integer.parseInt(ilvl); } catch (NumberFormatException ignored) {}
        }
    }

    private static Boolean toggle(Element el) {
        String v = attr(el, "val");
        return v == null || !(v.equals("0") || v.equalsIgnoreCase("false") || v.equalsIgnoreCase("off"));
    }

    // ------------------------------------------------------------------ numbering

    private static final class NumLevel {
        String fmt = "bullet";
        String text = "•";
        int start = 1;
        Float indLeft;
    }

    private void parseNumbering(Element root) {
        for (Element an : children(root, "abstractNum")) {
            String id = attr(an, "abstractNumId");
            if (id == null) continue;
            Map<Integer, NumLevel> levels = new HashMap<>();
            for (Element lvl : children(an, "lvl")) {
                NumLevel nl = new NumLevel();
                int il = parseInt(attr(lvl, "ilvl"), 0);
                String fmt = attr(child(lvl, "numFmt"), "val");
                if (fmt != null) nl.fmt = fmt;
                String txt = attr(child(lvl, "lvlText"), "val");
                if (txt != null) nl.text = txt;
                nl.start = parseInt(attr(child(lvl, "start"), "val"), 1);
                Element pPr = child(lvl, "pPr");
                if (pPr != null && child(pPr, "ind") != null) {
                    Element ind = child(pPr, "ind");
                    String left = attr(ind, "left");
                    if (left == null) left = attr(ind, "start");
                    if (left != null) nl.indLeft = twips(left, 0);
                }
                levels.put(il, nl);
            }
            abstractLevels.put(id, levels);
        }
        for (Element num : children(root, "num")) {
            String id = attr(num, "numId");
            String abs = attr(child(num, "abstractNumId"), "val");
            if (id != null && abs != null) numToAbstract.put(id, abs);
        }
    }

    private NumLevel numLevel(String numId, Integer ilvl) {
        if (numId == null) return null;
        String abs = numToAbstract.get(numId);
        if (abs == null) return null;
        Map<Integer, NumLevel> levels = abstractLevels.get(abs);
        if (levels == null) return null;
        return levels.get(ilvl == null ? 0 : ilvl);
    }

    private String listPrefix(Props eff) {
        if (eff.numId == null || "0".equals(eff.numId)) return null;
        int lvl = eff.ilvl == null ? 0 : Math.max(0, Math.min(eff.ilvl, 8));
        NumLevel level = numLevel(eff.numId, lvl);
        if (level == null) return "•";
        if ("bullet".equals(level.fmt)) {
            String t = level.text;
            // Symbol/Wingdings private-use glyphs don't render in normal fonts
            if (t == null || t.isEmpty() || t.charAt(0) >= '' || t.equals("o") || t.equals("§")) {
                return lvl % 3 == 1 ? "◦" : (lvl % 3 == 2 ? "▪" : "•");
            }
            return t;
        }
        if ("none".equals(level.fmt)) return null;

        int[] counters = numCounters.get(eff.numId);
        if (counters == null) {
            counters = new int[9];
            for (int k = 0; k < 9; k++) {
                NumLevel l = numLevel(eff.numId, k);
                counters[k] = (l != null ? l.start : 1) - 1;
            }
            numCounters.put(eff.numId, counters);
        }
        counters[lvl]++;
        for (int k = lvl + 1; k < 9; k++) {
            NumLevel l = numLevel(eff.numId, k);
            counters[k] = (l != null ? l.start : 1) - 1;
        }
        String text = level.text;
        for (int k = 0; k <= lvl; k++) {
            NumLevel l = numLevel(eff.numId, k);
            String fmt = l != null ? l.fmt : "decimal";
            text = text.replace("%" + (k + 1), formatNumber(Math.max(counters[k], 1), fmt));
        }
        return text;
    }

    private static String formatNumber(int n, String fmt) {
        switch (fmt) {
            case "lowerLetter": return letters(n).toLowerCase(Locale.ROOT);
            case "upperLetter": return letters(n);
            case "lowerRoman": return roman(n).toLowerCase(Locale.ROOT);
            case "upperRoman": return roman(n);
            default: return Integer.toString(n);
        }
    }

    private static String letters(int n) {
        StringBuilder sb = new StringBuilder();
        while (n > 0) { n--; sb.insert(0, (char) ('A' + n % 26)); n /= 26; }
        return sb.toString();
    }

    private static String roman(int n) {
        int[] v = {1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1};
        String[] s = {"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"};
        StringBuilder sb = new StringBuilder();
        for (int k = 0; k < v.length; k++) while (n >= v[k]) { n -= v[k]; sb.append(s[k]); }
        return sb.toString();
    }

    // ------------------------------------------------------------------ rels

    private void parseRels(Element root) {
        for (Element rel : children(root, "Relationship")) {
            String id = attr(rel, "Id");
            String target = attr(rel, "Target");
            if (id != null && target != null && !"External".equals(attr(rel, "TargetMode"))) relTargets.put(id, target);
        }
    }

    // ------------------------------------------------------------------ helpers

    private static Element xml(byte[] data) throws IOException {
        try {
            DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
            f.setNamespaceAware(true);
            try { f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true); } catch (Exception ignored) {}
            try { f.setExpandEntityReferences(false); } catch (Exception ignored) {}
            DocumentBuilder b = f.newDocumentBuilder();
            return b.parse(new ByteArrayInputStream(data)).getDocumentElement();
        } catch (Exception e) {
            throw new IOException("Could not read document XML: " + e.getMessage(), e);
        }
    }

    private static byte[] readAll(InputStream in) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[16384];
        int n;
        while ((n = in.read(buf)) > 0) bos.write(buf, 0, n);
        return bos.toByteArray();
    }

    static String local(Node n) {
        String l = n.getLocalName();
        if (l != null) return l;
        String name = n.getNodeName();
        int c = name.indexOf(':');
        return c >= 0 ? name.substring(c + 1) : name;
    }

    static List<Element> elements(Element parent) {
        List<Element> list = new ArrayList<>();
        if (parent == null) return list;
        NodeList nl = parent.getChildNodes();
        for (int k = 0; k < nl.getLength(); k++) {
            Node n = nl.item(k);
            if (n.getNodeType() == Node.ELEMENT_NODE) list.add((Element) n);
        }
        return list;
    }

    static Element child(Element parent, String localName) {
        if (parent == null) return null;
        NodeList nl = parent.getChildNodes();
        for (int k = 0; k < nl.getLength(); k++) {
            Node n = nl.item(k);
            if (n.getNodeType() == Node.ELEMENT_NODE && local(n).equals(localName)) return (Element) n;
        }
        return null;
    }

    static List<Element> children(Element parent, String localName) {
        List<Element> list = new ArrayList<>();
        for (Element e : elements(parent)) if (local(e).equals(localName)) list.add(e);
        return list;
    }

    static Element descendant(Element parent, String localName) {
        for (Element e : elements(parent)) {
            if (local(e).equals(localName)) return e;
            Element d = descendant(e, localName);
            if (d != null) return d;
        }
        return null;
    }

    /** Attribute value by local name (ignores namespace prefix). Returns null if absent. */
    static String attr(Element e, String localName) {
        if (e == null) return null;
        NamedNodeMap attrs = e.getAttributes();
        for (int k = 0; k < attrs.getLength(); k++) {
            Node a = attrs.item(k);
            if (local(a).equals(localName)) return a.getNodeValue();
        }
        return null;
    }

    private static float twips(String v, float def) {
        if (v == null) return def;
        try { return Float.parseFloat(v) / 20f; } catch (NumberFormatException e) { return def; }
    }

    private static float emu(String v) {
        if (v == null) return 0;
        try { return Float.parseFloat(v) / 12700f; } catch (NumberFormatException e) { return 0; }
    }

    private static int parseInt(String v, int def) {
        if (v == null) return def;
        try { return Integer.parseInt(v); } catch (NumberFormatException e) { return def; }
    }
}
