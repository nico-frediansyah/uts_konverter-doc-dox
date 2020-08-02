/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readwritedocx;

import java.io.File;
import org.apache.log4j.BasicConfigurator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
/**
 *
 * @author utrodus
 */
public class writeDocx {
    public static void main(String[] args) throws FileNotFoundException, IOException{
         BasicConfigurator.configure();
        
        String teks = "Java adalah sebuah bahasa pemrograman dasar dalam sebuah pembuatan aplikasi."
                + "Java juga merupakan bahasa pemrograman yang dapat di jalankan di berbagai komputer ataupun berbagai telepon genggam. "
                + "Kemudian, bahasa pemrograman java ini sendiri bisa digunakan untuk membuat sebuah game ataupun aplikasi untuk perangkat lunak maupun komputer sekalipun.\n" +
"";

        XWPFDocument documentDocx = new XWPFDocument();

        // write document
        String outDocxEn = "apaitujava.docx";
        //create paragraph
        try (FileOutputStream outDocx = new FileOutputStream(new File(outDocxEn))) {
            //create paragraph
            XWPFParagraph paragraphDocx = documentDocx.createParagraph();
            XWPFRun runDocx = paragraphDocx.createRun();
            runDocx.setText(teks);
            
            documentDocx.write(outDocx);
        }
        System.out.println("Write docx successfully");
    }
}
