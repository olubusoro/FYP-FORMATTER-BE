package com.example.demo;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import java.io.*;
import java.math.BigInteger;
import java.util.List;
import org.apache.poi.openxml4j.util.ZipSecureFile;

@RestController
@CrossOrigin(origins = "https://doc-formatter.vercel.app/")
public class FormatterController {

    @PostMapping("/format-my-project")
    public ResponseEntity<?> formatDoc(@RequestParam("file") MultipartFile file) {
        try {
            // 1. SECURITY
            ZipSecureFile.setMinInflateRatio(0.001);

            if (file.isEmpty()) return ResponseEntity.badRequest().body("File is empty.");
            if (!file.getOriginalFilename().endsWith(".docx")) return ResponseEntity.badRequest().body("Upload .docx only.");

            XWPFDocument document = new XWPFDocument(file.getInputStream());

            // 2. CLEANUP (Standard)
            removeExistingHeadersFooters(document);

            // 3. GHOSTBUSTER: Delete empty "Enters" above chapters to stop blank pages
            cleanUpGhostLines(document);

            // 4. LAYOUT
            setProfessionalMargins(document);

            // 5. THE ENGINE
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int startProcessingIndex = 0;

            // START LOGIC
            for(int i = 0; i < paragraphs.size(); i++) {
                String t = paragraphs.get(i).getText().trim().toUpperCase();
                if(t.startsWith("CHAPTER") || t.startsWith("ABSTRACT") || t.startsWith("INTRODUCTION")
                        || t.startsWith("DEDICATION") || t.startsWith("ACKNOWLEDGEMENT")) {
                    startProcessingIndex = i;
                    break;
                }
            }

            boolean nextLineIsChapterTitle = false;

            for (int i = startProcessingIndex; i < paragraphs.size(); i++) {
                XWPFParagraph p = paragraphs.get(i);

                // Double Spacing
                p.setSpacingLineRule(LineSpacingRule.AUTO);
                p.setSpacingBetween(2.0);

                String text = p.getText().trim();
                String upperText = text.toUpperCase();

                if (text.isEmpty()) {
                    nextLineIsChapterTitle = false;
                    continue;
                }

                // IDENTIFY TYPES
                boolean isChapter = upperText.startsWith("CHAPTER");
                boolean isMajorSection = false;
                if (upperText.startsWith("DEDICATION")) isMajorSection = true;
                if (upperText.startsWith("ACKNOWLEDGEMENT")) isMajorSection = true;
                if (upperText.startsWith("ABSTRACT")) isMajorSection = true;
                if (upperText.startsWith("REFERENCES")) isMajorSection = true;

                // *** FIX: STRICT REGEX (Must have a DOT to be a heading) ***
                // Matches "1.1", "2.3.1" but IGNORES "2008", "1999"
                boolean isSubHeading = text.matches("^[0-9]+\\.[0-9]+.*");

                if (isChapter || isMajorSection) {
                    // Safety Clean: Remove manual breaks
                    removeManualPageBreaks(p);
                    if (i > 0) {
                        XWPFParagraph prev = paragraphs.get(i - 1);
                        removeManualPageBreaks(prev);
                        prev.setPageBreak(false);
                    }

                    p.setAlignment(ParagraphAlignment.CENTER);
                    p.setStyle("Heading 1");
                    p.setPageBreak(true); // Force 1 Clean New Page

                    if (isChapter) nextLineIsChapterTitle = true;

                    applyFont(p, true, 14);
                }
                else if (nextLineIsChapterTitle) {
                    p.setAlignment(ParagraphAlignment.CENTER);
                    p.setStyle("Heading 1");
                    p.setPageBreak(false);

                    applyFont(p, true, 14);
                    nextLineIsChapterTitle = false;
                }
                else if (isSubHeading) {
                    p.setAlignment(ParagraphAlignment.BOTH);
                    p.setStyle("Heading 2");
                    p.setPageBreak(false);
                    applyFont(p, true, 12);
                }
                else {
                    p.setAlignment(ParagraphAlignment.BOTH);
                    p.setPageBreak(false);
                    applyFont(p, false, 12);
                }
            }

            // 6. PAGE NUMBERS
            addAggressivePageNumbers(document);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            document.write(out);
            document.close();

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Formatted_Project.docx")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(out.toByteArray());

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body("Error: " + e.getMessage());
        }
    }

    // --- HELPERS ---

    private void cleanUpGhostLines(XWPFDocument document) {
        // Iterate backwards to safely delete lines
        List<XWPFParagraph> paras = document.getParagraphs();
        for (int i = paras.size() - 1; i >= 0; i--) {
            String text = paras.get(i).getText().trim().toUpperCase();

            // If we find a Heading
            if (text.startsWith("CHAPTER") || text.startsWith("ABSTRACT") ||
                    text.startsWith("DEDICATION") || text.startsWith("ACKNOWLEDGEMENT") ||
                    text.startsWith("REFERENCES")) {

                // Look ABOVE it and delete all empty lines
                int j = i - 1;
                while (j >= 0) {
                    XWPFParagraph prev = paras.get(j);
                    if (prev.getText().trim().isEmpty()) {
                        int pos = document.getPosOfParagraph(prev);
                        if(pos != -1) document.removeBodyElement(pos);
                        j--;
                    } else {
                        break;
                    }
                }
            }
        }
    }

    private void removeManualPageBreaks(XWPFParagraph p) {
        for (XWPFRun run : p.getRuns()) {
            List<CTBr> brList = run.getCTR().getBrList();
            if (brList != null && !brList.isEmpty()) {
                run.getCTR().setBrArray(new CTBr[0]);
            }
        }
    }

    private void applyFont(XWPFParagraph p, boolean isBold, int fontSize) {
        for (XWPFRun run : p.getRuns()) {
            run.setFontFamily("Times New Roman");
            run.setFontSize(fontSize);
            run.setBold(isBold);
        }
    }

    private void addAggressivePageNumbers(XWPFDocument document) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph p = footer.createParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun r1 = p.createRun(); r1.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        XWPFRun r2 = p.createRun(); r2.getCTR().addNewInstrText().setStringValue("PAGE");
        XWPFRun r3 = p.createRun(); r3.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        XWPFRun r4 = p.createRun(); r4.setText("1");
        XWPFRun r5 = p.createRun(); r5.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
        linkFooterToSection(document, sectPr, footer);

        for (XWPFParagraph para : document.getParagraphs()) {
            if (para.getCTP().getPPr() != null && para.getCTP().getPPr().getSectPr() != null) {
                linkFooterToSection(document, para.getCTP().getPPr().getSectPr(), footer);
            }
        }
    }

    private void linkFooterToSection(XWPFDocument document, CTSectPr sectPr, XWPFFooter footer) {
        if (sectPr.isSetTitlePg()) sectPr.unsetTitlePg();
        sectPr.getFooterReferenceList().clear();
        CTHdrFtrRef ref = sectPr.addNewFooterReference();
        ref.setType(STHdrFtr.DEFAULT);
        ref.setId(document.getRelationId(footer));
    }

    private void removeExistingHeadersFooters(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr != null) {
            sectPr.getFooterReferenceList().clear();
            sectPr.getHeaderReferenceList().clear();
        }
        for (XWPFFooter footer : document.getFooterList()) {
            while (footer.getParagraphs().size() > 0) {
                footer.removeParagraph(footer.getParagraphs().get(0));
            }
        }
    }

    private void setProfessionalMargins(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(2160));
        pageMar.setRight(BigInteger.valueOf(1440));
        pageMar.setTop(BigInteger.valueOf(1440));
        pageMar.setBottom(BigInteger.valueOf(1440));
    }
}