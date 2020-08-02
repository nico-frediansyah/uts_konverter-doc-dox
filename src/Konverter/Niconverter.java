/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Konverter;

import java.awt.HeadlessException;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author 62813
 */
public class Niconverter extends javax.swing.JFrame {

    /**
     * Creates new form Niconverter
     */
    public Niconverter() {
        initComponents();
        this.setTitle("Aplikasi Converter File");
        this.setLocationRelativeTo(this);
        docKonvers.setEnabled(false);
        docxKonvers.setEnabled(false);
    }

    public void konvertKeDoc(String path) {
        JFileChooser chooser = new JFileChooser(".");
        chooser.setFileFilter(new FileNameExtensionFilter(".doc", "doc"));
        int buka_dialog = chooser.showSaveDialog(Niconverter.this);
        if (buka_dialog == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().toString();
            if (!filePath.endsWith(".doc")) {
                filePath += ".doc";
            }
            fileOutput.setText(filePath);

            String line = null;
            ArrayList textFile = new ArrayList();
            try {
                // Baca File Txt
                FileReader fileReader = new FileReader(path);
                // membaca input file / isi file
                BufferedReader bufferedReader = new BufferedReader(fileReader);
                while ((line = bufferedReader.readLine()) != null) {
                    textFile.add(line);
                }
                bufferedReader.close();
                writeKeDoc(filePath, textFile);

            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "File Tidak Ada");
            }

        }
    }

    public void konvertKeDocx(String path) {
        JFileChooser chooser = new JFileChooser(".");
        chooser.setFileFilter(new FileNameExtensionFilter(".docx", "docx"));
        int buka_dialog = chooser.showSaveDialog(Niconverter.this);
        if (buka_dialog == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().toString();
            if (!filePath.endsWith(".docx")) {
                filePath += ".docx";
            }
            fileOutput.setText(filePath);

            // Baca File Txt
            String line = null;
            ArrayList textFile = new ArrayList();
            try {
                FileReader fileReader = new FileReader(path);
                // membaca input file / isi file
                BufferedReader bufferedReader = new BufferedReader(fileReader);
                while ((line = bufferedReader.readLine()) != null) {
                    textFile.add(line);
                }
                bufferedReader.close();
                writeKeDocx(filePath, textFile);

            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "File Tidak Ada");
            }

        }
    }

    public void writeKeDoc(String filePath, ArrayList textFile) {
        try {

            Properties prop = new Properties();
            prop.setProperty("log4j.rootLogger", "WARN");

            // membuat dokumen
            String outDocEn = filePath;
            XWPFDocument document = new XWPFDocument();

            // membuat file
            FileOutputStream out = new FileOutputStream(new File(outDocEn));

            // membuat paragraf
            for (int i = 0; i < textFile.size(); i++) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(String.valueOf(textFile.get(i)));
            }

            document.write(out);
            out.close();

            JOptionPane.showMessageDialog(null, "Convert ke Doc Berhasil");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(e);
        }
    }

    public void writeKeDocx(String filePath, ArrayList textFile) {
        try {

            Properties prop = new Properties();
            prop.setProperty("log4j.rootLogger", "WARN");

            // membuat dokumen
            String outDocEn = filePath;
            XWPFDocument document = new XWPFDocument();

            // membuat paragraf
            try ( // membuat file
                    FileOutputStream out = new FileOutputStream(new File(outDocEn))) {
                // membuat paragraf
                for (int i = 0; i < textFile.size(); i++) {
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();
                    run.setText(String.valueOf(textFile.get(i)));
                }
                
                document.write(out);
            }

            JOptionPane.showMessageDialog(null, "Konversi ke Docx Berhasil");
        } catch (HeadlessException | IOException e) {
            System.out.println(e);
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        fileOutput = new javax.swing.JTextField();
        docxKonvers = new javax.swing.JButton();
        docKonvers = new javax.swing.JButton();
        fileInput = new javax.swing.JButton();
        pathFileAwal = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setBackground(new java.awt.Color(0, 204, 204));

        jPanel1.setBackground(new java.awt.Color(153, 153, 255));
        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder(null, " \" Converter txt file to *.doc/*.docx \" ", javax.swing.border.TitledBorder.CENTER, javax.swing.border.TitledBorder.TOP, new java.awt.Font("Courier New", 1, 14), new java.awt.Color(255, 255, 255))); // NOI18N
        jPanel1.setForeground(new java.awt.Color(102, 102, 255));

        docxKonvers.setBackground(new java.awt.Color(0, 255, 0));
        docxKonvers.setFont(new java.awt.Font("Serif", 1, 11)); // NOI18N
        docxKonvers.setText("*.docx");
        docxKonvers.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                docxKonversActionPerformed(evt);
            }
        });

        docKonvers.setBackground(new java.awt.Color(255, 0, 51));
        docKonvers.setFont(new java.awt.Font("Serif", 1, 12)); // NOI18N
        docKonvers.setText("*.doc");
        docKonvers.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                docKonversActionPerformed(evt);
            }
        });

        fileInput.setText("Open File");
        fileInput.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileInputActionPerformed(evt);
            }
        });

        jLabel1.setBackground(new java.awt.Color(255, 255, 255));
        jLabel1.setFont(new java.awt.Font("Traditional Arabic", 0, 18)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Convert file to...");

        jLabel2.setFont(new java.awt.Font("Verdana", 0, 11)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 204));
        jLabel2.setText("File output :");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(pathFileAwal, javax.swing.GroupLayout.PREFERRED_SIZE, 230, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(fileInput))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(0, 5, Short.MAX_VALUE)
                        .addComponent(jLabel2)
                        .addGap(18, 18, 18)
                        .addComponent(fileOutput, javax.swing.GroupLayout.PREFERRED_SIZE, 230, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(48, 48, 48)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(48, 48, 48)
                        .addComponent(jLabel1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 72, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addComponent(docKonvers, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(38, 38, 38)
                        .addComponent(docxKonvers, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(fileInput)
                    .addComponent(pathFileAwal, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(5, 5, 5)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(docxKonvers)
                    .addComponent(docKonvers))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(fileOutput, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void fileInputActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_fileInputActionPerformed
        // TODO add your handling code here:
        File filenya;
        JFileChooser chooser = new JFileChooser(".");
        chooser.setFileFilter(new FileNameExtensionFilter(".txt", "txt"));
        int buka_dialog = chooser.showOpenDialog(Niconverter.this);
        if (buka_dialog == JFileChooser.APPROVE_OPTION) {
            filenya = chooser.getSelectedFile();
            String filePath = filenya.getPath();
            String fileName = filenya.getName();
            try {
                String fileExtention = fileName.substring(fileName.lastIndexOf("."), fileName.length());
                if (!".txt".equals(fileExtention)) {
                    JOptionPane.showMessageDialog(null, "Maaf ! Hanya dapat menerima Format File .txt");
                } else {
                    pathFileAwal.setText(filePath);

                    if (".txt".equals(fileExtention)) {
                        docKonvers.setEnabled(true);
                        docxKonvers.setEnabled(true);
                    }
                }
            } catch (Exception e) {
                System.out.println(e);
//                JOptionPane.showMessageDialog(null, "Maaf ! Hanya dapat menerima Format File .txt atau .doc ");
            }

        }
    }//GEN-LAST:event_fileInputActionPerformed

    private void docKonversActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_docKonversActionPerformed
        // TODO add your handling code here:
        konvertKeDoc(pathFileAwal.getText());
    }//GEN-LAST:event_docKonversActionPerformed

    private void docxKonversActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_docxKonversActionPerformed
        // TODO add your handling code here:
        konvertKeDocx(pathFileAwal.getText());
    }//GEN-LAST:event_docxKonversActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Classic".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Niconverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Niconverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Niconverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Niconverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Niconverter().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton docKonvers;
    private javax.swing.JButton docxKonvers;
    private javax.swing.JButton fileInput;
    private javax.swing.JTextField fileOutput;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField pathFileAwal;
    // End of variables declaration//GEN-END:variables
}
