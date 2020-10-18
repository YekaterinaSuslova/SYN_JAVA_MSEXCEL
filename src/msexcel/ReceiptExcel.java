package msexcel;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReceiptExcel extends javax.swing.JFrame {

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий каталог
            try {
                modifData(dir + "receipt_template.xls", dir + "receipt.xls", jTextField1_Name.getText(),
                        jTextField2_Code.getText(), jTextField3_Position.getText(),
                        jTextField4_Department.getText()); // Вызов метода создания отчета
                Desktop.getDesktop().open(new File(dir + "receipt.xls")); // Запуск отчета в MS Excel
            } catch (Exception ex) {
                System.err.println("Error modifData!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    public ReceiptExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField1_Name = new javax.swing.JTextField();
        jTextField2_Code = new javax.swing.JTextField();
        jTextField3_Position = new javax.swing.JTextField();
        jTextField4_Department = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Квитанция в MS Excel");
        getContentPane().setLayout(null);
        getContentPane().add(jTextField1_Name);
        jTextField1_Name.setBounds(120, 100, 180, 20);
        getContentPane().add(jTextField2_Code);
        jTextField2_Code.setBounds(120, 120, 180, 20);
        getContentPane().add(jTextField3_Position);
        jTextField3_Position.setBounds(120, 140, 180, 20);
        getContentPane().add(jTextField4_Department);
        jTextField4_Department.setBounds(400, 100, 190, 20);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(860, 400, 90, 23);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/msexcel/receipt.png"))); // NOI18N
        jLabel1.setText("jLabel1");
        getContentPane().add(jLabel1);
        jLabel1.setBounds(20, 10, 1070, 450);

        setSize(new java.awt.Dimension(1150, 507));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed

        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();

    }//GEN-LAST:event_jButton1ActionPerformed

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptExcel().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField1_Name;
    private javax.swing.JTextField jTextField2_Code;
    private javax.swing.JTextField jTextField3_Position;
    private javax.swing.JTextField jTextField4_Department;
    // End of variables declaration//GEN-END:variables
private void modifData(String inputFileName, String outputFileName, String Name, String Code,
            String Position, String Department) throws IOException {
        // Метод создания отчета
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(inputFileName)); // Файл-шаблон MS Excel
        HSSFWorkbook wb = new HSSFWorkbook(fs); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel
        HSSFRow row; // Строка
        HSSFCell cell; // Ячейка

        row = sheet.getRow(3); // Выбираем строку
        cell = row.getCell(1);  // Выбираем столбец стоки
        cell.setCellValue(Name); // Устанавливаем значение ячейки [D13]: (4,13)

        row = sheet.getRow(4);
        cell = row.getCell(1);
        cell.setCellValue(Code);
        
        row = sheet.getRow(5);
        cell = row.getCell(1);
        cell.setCellValue(Position);
        
        row = sheet.getRow(3);
        cell = row.getCell(4);
        cell.setCellValue(Department);

        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

}
