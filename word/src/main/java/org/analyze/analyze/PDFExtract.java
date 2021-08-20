package org.analyze.analyze;

import org.apache.pdfbox.io.RandomAccessBuffer;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineNode;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

/**
 * @Author: zyj
 * @Date: 2021/7/1
 */
@Component
public class PDFExtract {
    public void extract(MultipartFile file){

        try(InputStream inputStream = file.getInputStream()){

            PDDocument document = null;
            PDFParser parser = new PDFParser(new RandomAccessBuffer(inputStream));
            parser.parse();
            document = parser.getPDDocument();

            PDDocumentOutline outline =  document.getDocumentCatalog().getDocumentOutline();

            if( outline != null ) {
                printBookmark( outline, "" );
            } else {
                System.out.println( "This document does not contain any bookmarks" );
            }
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    //读取文档书签
    public void printBookmark(PDOutlineNode bookmark, String indentation ) throws IOException {
        PDOutlineItem current = bookmark.getFirstChild();
        while( current != null ) {
            int pages =0;
            if (current.getDestination() instanceof PDPageDestination)
            {
                PDPageDestination pd = (PDPageDestination) current.getDestination();
                pages = (pd.retrievePageNumber() +1);
            }
            if (current.getAction() instanceof PDActionGoTo)
            {
                PDActionGoTo gta = (PDActionGoTo) current.getAction();
                if (gta.getDestination() instanceof PDPageDestination)
                {
                    PDPageDestination pd = (PDPageDestination) gta.getDestination();
                    pages = (pd.retrievePageNumber() +1);
                }
            }
            if (pages ==0)
                System.out.println( "   " +indentation  + current.getTitle());
            else
                System.out.println( "   " +indentation + current.getTitle() +"  "+ pages);
            printBookmark( current,  "   "  + indentation  );  // 递归调用
            current = current.getNextSibling();
        }

    }

    public void printText(){

    }

}
