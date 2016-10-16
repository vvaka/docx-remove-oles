package com.poitest;

/**
 * Created by vvaka on 10/15/16.
 */

import org.apache.poi.POIOLE2TextExtractor;
import org.apache.poi.POITextExtractor;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.hssf.usermodel.HSSFObjectData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

public class POIMain {

    public static void readDocFile(String fileName) {

        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            HWPFDocument doc = new HWPFDocument(fis);
            WordExtractor we = new WordExtractor(doc);
            String[] paragraphs = we.getParagraphText();
            System.out.println("Total no of paragraph " + paragraphs.length);
            for (String para : paragraphs) {
                System.out.println(para.toString());
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void listEmbeds(XWPFDocument doc) throws OpenXML4JException {
        List<PackagePart> embeddedDocs = doc.getAllEmbedds();
        if (embeddedDocs != null && !embeddedDocs.isEmpty()) {
            Iterator<PackagePart> pIter = embeddedDocs.iterator();
            while (pIter.hasNext()) {
                PackagePart pPart = pIter.next();
                System.out.print(pPart.getContentTypeDetails().getParameterKeys() + ", ");
                System.out.print(pPart.getContentType() + ", ");
                System.out.println();
            }
        }
    }

    private static XWPFDocument removeEmbeds(XWPFDocument doc) throws OpenXML4JException {

        Iterator i$ = doc.getPackagePart().getRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject").iterator();

        PackageRelationship rel;
        while (i$.hasNext()) {
            i$.next();
            i$.remove();
        }

        i$ = doc.getPackagePart().getRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package").iterator();

        while (i$.hasNext()) {
            i$.next();
            i$.remove();

        }

        return doc;
    }

    public static void modifyDocxFile(String fileName) {
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            System.out.println(document.getAllEmbedds().size());

            removeEmbeds(document);
            System.out.println(document.getAllEmbedds());
            // embeds removed, save the file

            FileOutputStream out = new FileOutputStream("simpleM.docx");

            document.write(out);

            out.close();

            //System.out.println(document.getAllEmbedds());
            fis.close();
            //fis2.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        try {
            modifyDocxFile("Sample2.docx");
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        //readDocFile("C:\\Test.doc");
    }

    public static void ReadCSV(String fileName) throws IOException {
        FileInputStream myInput = new FileInputStream(fileName);
        POIFSFileSystem fs = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(fs);
        for (HSSFObjectData obj : workbook.getAllEmbeddedObjects()) {
            //the OLE2 Class Name of the object
            System.out.println("Objects : " + obj.getOLE2ClassName() + "   2 .");
            String oleName = obj.getOLE2ClassName();
            if (oleName.equals("Worksheet")) {
                // some code to process embedded excel file;
            } else if (oleName.equals("Document")) {
                System.out.println("Document");
                DirectoryNode dn = (DirectoryNode) obj.getDirectory();
                HWPFDocument embeddedWordDocument = new HWPFDocument(dn);
                System.out.println("Doc : " + embeddedWordDocument.getRange().text());
                // want to extract document not text into a doc file
                //************************
                //FileOutputStream fos = new FileOutputStream("E:\\log.txt");
                //fos.write(text.getBytes());
                //************************
            } else if (oleName.equals("Presentation")) {
                // some code to process embedded power point file;
            } else {
                // some code to process other kind of embedded files;
            }
        }
    }

    public static void getDocs(String inputFile) throws Exception {
        FileInputStream fis = new FileInputStream(inputFile);
        POIFSFileSystem fileSystem = new POIFSFileSystem(fis);
        // Firstly, get an extractor for the Workbook
        POIOLE2TextExtractor oleTextExtractor =
                ExtractorFactory.createExtractor(fileSystem);
        // Then a List of extractors for any embedded Excel, Word, PowerPoint
        // or Visio objects embedded into it.
        POITextExtractor[] embeddedExtractors = ExtractorFactory.getEmbededDocsTextExtractors(oleTextExtractor);
        for (POITextExtractor textExtractor : embeddedExtractors) {
            // A Word Document
            if (textExtractor instanceof WordExtractor) {
                WordExtractor wordExtractor = (WordExtractor) textExtractor;
                String[] paragraphText = wordExtractor.getParagraphText();
                for (String paragraph : paragraphText) {
                    System.out.println(paragraph);
                }
                // Display the document's header and footer text
                System.out.println("Footer text: " + wordExtractor.getFooterText());
                System.out.println("Header text: " + wordExtractor.getHeaderText());
            }

        }

    }
}