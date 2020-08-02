/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readwritedoc;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

/**
 *
 * @author 62813
 */
public class ReadDoc {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        File file = new File("D:readDoc.doc");
        WordExtractor extractor = null;
        try {
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            HWPFDocument dokumen = new HWPFDocument(fis);
            extractor = new WordExtractor(dokumen);
            String teks = extractor.getText();
            System.out.println(teks);
        } catch (Exception exep) {
            System.out.println("ex");
        }

    }

}
