/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.russianpost.siberia.monitorticket;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;
import java.util.regex.Pattern;
import javax.swing.JFileChooser;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.netbeans.api.progress.ProgressHandle;
import org.netbeans.api.settings.ConvertAsProperties;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.util.Exceptions;
import org.openide.windows.TopComponent;
import org.openide.util.NbBundle.Messages;
import ru.russianpost.siberia.maveneeticketlibrary.api.FindTicket;
import ru.russianpost.siberia.maveneeticketlibrary.api.FindTicketResponse;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketSRV;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketSRV_Service;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTickets;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketsResponse;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetReadyAnswer;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetReadyAnswerResponse;
import ru.russianpost.siberia.maveneeticketlibrary.api.Historyrecord;
import ru.russianpost.siberia.maveneeticketlibrary.api.Ticket;
import ru.russianpost.siberia.maveneeticketlibrary.api.ViewHistorySERV;
import ru.russianpost.siberia.maveneeticketlibrary.api.ViewHistorySERV_Service;

/**
 * Top component which displays something.
 */
@ConvertAsProperties(
        dtd = "-//ru.russianpost.siberia.monitorticket//batchmonitorform//EN",
        autostore = false
)
@TopComponent.Description(
        preferredID = "batchmonitorformTopComponent",
        //iconBase="SET/PATH/TO/ICON/HERE", 
        persistenceType = TopComponent.PERSISTENCE_ALWAYS
)
@TopComponent.Registration(mode = "editor", openAtStartup = false)
@ActionID(category = "Window", id = "ru.russianpost.siberia.monitorticket.batchmonitorformTopComponent")
@ActionReference(path = "Menu/Window" /*, position = 333 */)
@TopComponent.OpenActionRegistration(
        displayName = "#CTL_batchmonitorformAction",
        preferredID = "batchmonitorformTopComponent"
)
@Messages({
    "CTL_batchmonitorformAction=Пакетный поиск",
    "CTL_batchmonitorformTopComponent=Пакетный поиск РПО",
    "HINT_batchmonitorformTopComponent=This is a batchmonitorform window"
})
public final class batchmonitorformTopComponent extends TopComponent {

    File reqfile;
    List<String> lines;

    public batchmonitorformTopComponent() {
        lines = new ArrayList<>();
        initComponents();
        setName(Bundle.CTL_batchmonitorformTopComponent());
        setToolTipText(Bundle.HINT_batchmonitorformTopComponent());
    }

    /*
    Читаем из файла данные и записываем в таблицу ticket
     */
    private List<String> readFromFileTickets(List<String> lines, File file) {
        BufferedReader br = null;
        String line;
        String cvsSplitBy = "|";
        try {
//            br = new BufferedReader(new FileReader(file));
            br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
            while ((line = br.readLine()) != null) {
                if (!validateTicketFormat(lines, line)) {
                    String[] data = line.split(Pattern.quote(cvsSplitBy));
                    validateTicketFormat(lines, data[0]);
                }
            }
//            lines = Files.readAllLines(file.toPath(), StandardCharsets.UTF_8);
            util.LogDebug("Read - " + String.valueOf(lines.size()) + " barcode");
        } catch (IOException ex) {
            util.LogErr(ex.getMessage());
        } finally {
            if (br != null) {
                try {
                    br.close();
                } catch (IOException ex) {
                    util.LogErr(ex.getMessage());
                }
            }
        }
        return lines;
    }
    // заполнение строки (rowNum) определенного листа (sheet)
    // данными  из dataModel созданного в памяти Excel файла

    private int createSheetHeader(HSSFSheet sheet, int rowNum, Ticket t) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(t.getBarcode());
        row.createCell(1).setCellValue(t.getDateFetch().toString());
        row.createCell(2).setCellValue(t.isIsFinal());
        for (Historyrecord h : t.getHistoryrecord()) {
            Row r = sheet.createRow(++rowNum);
            r.createCell(1).setCellValue(h.getOperationAddressIndex());
            r.createCell(2).setCellValue(h.getOperDate().toString());
            r.createCell(3).setCellValue(h.getOperTypeID());
            r.createCell(4).setCellValue(h.getOperTypeName());
            r.createCell(5).setCellValue(h.getOperAttrID());
            r.createCell(6).setCellValue(h.getOperAttrName());
            r.createCell(7).setCellValue(h.getOperatonDelta());
        }
        return rowNum;
    }

    private boolean validateTicketFormat(List<String> lines, String barcode) {
        if ((barcode.length() == 14) || (barcode.length() == 13)) {
            lines.add(barcode.toUpperCase());
            return true;
        }
        return false;
    }

    private File GetFileRequest(boolean isSave, FileNameExtensionFilter filter) {
        JFileChooser saveFile = new JFileChooser();
        saveFile.setFileFilter(filter);
        if (isSave) {
            if (saveFile.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
                return saveFile.getSelectedFile();
            }
        } else {
            if (saveFile.showDialog(null, "Выберите файл") == JFileChooser.APPROVE_OPTION) {
                return saveFile.getSelectedFile();
            }
        }
        return null;
    }

    private List<String> readFromExcelTickets(List<String> lines, File file) {
        // создание самого excel файла в памяти
        org.apache.poi.ss.usermodel.Workbook workbook;
        try {
            workbook = WorkbookFactory.create(file);
            // Get first sheet from the workbook
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
            // Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = sheet.iterator();
            DataFormatter formatter = new DataFormatter(); //creating formatter
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // Get iterator to all cells of current row 
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    // Change to getCellType() if using POI 4.x
                    CellType cellType = cell.getCellTypeEnum();
                    switch (cellType) {
//                        case NUMERIC:
//                            validateTicketFormat(lines, formatter.formatCellValue(cell));
//                            break;
//                        case STRING:
//                            validateTicketFormat(lines, formatter.formatCellValue(cell));
//                            break;
                        case FORMULA:
                            validateTicketFormat(lines, formatter.formatCellValue(cell, evaluator));
                            break;
                        default:
                            validateTicketFormat(lines, formatter.formatCellValue(cell));
                    }
                }
            }
        } catch (InvalidFormatException | EncryptedDocumentException | IOException ex) {
            util.LogErr(ex.getMessage());
        }
        return lines;
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        btFileLoad = new javax.swing.JButton();
        btLoadAnswer = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        InfoTrace = new javax.swing.JTextArea();

        setNextFocusableComponent(btFileLoad);

        org.openide.awt.Mnemonics.setLocalizedText(btFileLoad, org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btFileLoad.text")); // NOI18N
        btFileLoad.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btFileLoadActionPerformed(evt);
            }
        });

        org.openide.awt.Mnemonics.setLocalizedText(btLoadAnswer, org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btLoadAnswer.text")); // NOI18N
        btLoadAnswer.setToolTipText(org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btLoadAnswer.toolTipText")); // NOI18N
        btLoadAnswer.setEnabled(false);
        btLoadAnswer.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btLoadAnswerActionPerformed(evt);
            }
        });

        InfoTrace.setColumns(20);
        InfoTrace.setRows(5);
        jScrollPane2.setViewportView(InfoTrace);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btFileLoad, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btLoadAnswer, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 374, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(18, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(16, 16, 16)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 129, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(btFileLoad)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(btLoadAnswer)))
                .addContainerGap(153, Short.MAX_VALUE))
        );
    }// </editor-fold>//GEN-END:initComponents

    private class WaittoAnswer extends SwingWorker<Void, String> {

        private final List<String> request;
        private final GetBatchTicketSRV_Service service;
        private final GetBatchTicketSRV port;
        private final GetReadyAnswer parameters;
        long start;

        public WaittoAnswer(List<String> reqticket) {
            request = reqticket;
            service = new GetBatchTicketSRV_Service();
            port = service.getGetBatchTicketSRVPort();
            parameters = new GetReadyAnswer();
        }

        @Override
        protected Void doInBackground() throws Exception {
            // запоминаем текущее время в миллисекундах 
            start = System.currentTimeMillis();
            boolean f = true;
            while (f) {
                Thread.sleep(1000L); //One minute
                f = false;
                for (String req : request) {
                    parameters.setReq(req);
                    GetReadyAnswerResponse r = port.getReadyAnswer(parameters);
                    f = f | r.isReturn();
                    publish(".");
                }
            }
            return null;
        }

        @Override
        protected void done() {
            btLoadAnswer.setText("Получить");
            btLoadAnswer.setEnabled(true);
            InfoTrace.append("Ok!\n");
        }

        @Override
        protected void process(List<String> chunks) {
            for (String chunk : chunks) {
                btLoadAnswer.setText(String.valueOf(System.currentTimeMillis() - start));
                InfoTrace.append(chunk);
            }
        }

    }

    private void btFileLoadActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btFileLoadActionPerformed
        util.changeCursorWaitStatus(true);
        btFileLoad.setEnabled(false);
        if ((reqfile = GetFileRequest(false, new FileNameExtensionFilter("TXT & CVS & XLS Files", "txt", "cvs", "xls"))) != null) {
            try {
                String filename = reqfile.getName().toLowerCase();
                InfoTrace.append("Загружаем фйал с ШПИ: " + filename + "\n");
                util.LogDebug("File name - " + filename);
                lines.clear();
                if ((filename.indexOf("txt") > 0) | (filename.indexOf("cvs") > 0)) {
                    InfoTrace.append("Определен текстовый формат\n");
                    lines = readFromFileTickets(lines, reqfile);
                } else if (filename.indexOf("xls") > 0) {
                    InfoTrace.append("Определен excel формат\n");
                    lines = readFromExcelTickets(lines, reqfile);
                } else {
                    InfoTrace.append("Формат не определен\n");
                    InfoTrace.append("Завершено\n");
                    btFileLoad.setEnabled(true);
                }
                if (lines.size() > 0) {
                    List<String> request = GetBatchTicketService(lines);
                    InfoTrace.append("Ожидаем ответа от сервера");
                    WaittoAnswer task = new WaittoAnswer(request);
                    task.execute();
                }
            } catch (Exception ex) {
                Exceptions.printStackTrace(ex);
            }
        }
        util.changeCursorWaitStatus(false);
    }//GEN-LAST:event_btFileLoadActionPerformed

    private void btLoadAnswerActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btLoadAnswerActionPerformed
        GetAnswerServer(lines);
        btFileLoad.setEnabled(true);
        btLoadAnswer.setEnabled(false);
        lines.clear();
    }//GEN-LAST:event_btLoadAnswerActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextArea InfoTrace;
    private javax.swing.JButton btFileLoad;
    private javax.swing.JButton btLoadAnswer;
    private javax.swing.JScrollPane jScrollPane2;
    // End of variables declaration//GEN-END:variables
    @Override
    public void componentOpened() {
    }

    @Override
    public void componentClosed() {
    }

    void writeProperties(java.util.Properties p) {
        p.setProperty("version", "1.0");
    }

    void readProperties(java.util.Properties p) {
        String version = p.getProperty("version");
    }

    /*
    Отправляем данные на сервер
     */
    private List<String> GetBatchTicketService(List<String> lines) {
        List<String> result = null;
        try { // Call Web Service Operation
            InfoTrace.append("Отправляем запрос из " + String.valueOf(lines.size()) + " ШПИ на сервер...");
            GetBatchTicketSRV_Service service = new GetBatchTicketSRV_Service();
            GetBatchTicketSRV port = service.getGetBatchTicketSRVPort();
            GetBatchTickets parameters = new GetBatchTickets();
            parameters.getTickets().addAll(lines);
            GetBatchTicketsResponse resp = port.getBatchTickets(parameters);
            InfoTrace.append("отправили\n");
            util.LogDebug("Get server done.");
            result = resp.getReturn();
        } catch (Exception ex) {
            InfoTrace.append("ошибка\n");
            util.LogErr(ex.getMessage());
        }
        return result;
    }

    /*
    Получаем данные с сервера
     */
    private void GetAnswerServer(List<String> lines) {
        File saveFile = GetFileRequest(true, new FileNameExtensionFilter("XLS Files", "xls"));
        InfoTrace.append("Выгружаем в файл " + saveFile.getName() + "\n");
        if (saveFile != null) {
            class Task implements Runnable {

                List<String> lines;
                File saveFile;

                Task(File f, List<String> l) {
                    saveFile = f;
                    lines = new ArrayList<>();
                    lines.addAll(l);
                }

                public void run() {
                    ProgressHandle p = ProgressHandle.createHandle("Загрузка ШПИ с сервера");
                    p.start(lines.size());
                    // создание самого excel файла в памяти
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    // создание листа с названием "Просто лист"
                    HSSFSheet sheet = workbook.createSheet("ШПИ");
                    // счетчик для строк
                    int rowNum = 0;
                    // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
                    Row row = sheet.createRow(rowNum);
                    row.createCell(0).setCellValue("ШПИ");
                    row.createCell(1).setCellValue("Дата");
                    row.createCell(2).setCellValue("Финал");
                    ViewHistorySERV_Service service = new ViewHistorySERV_Service();
                    ViewHistorySERV port = service.getViewHistorySERVPort();
                    FindTicket parameters = new FindTicket();
                    int i = 1;
                    for (String line : lines) {
                        try { // Call Web Service Operation
                            parameters.setBarcode(line);
                            FindTicketResponse tk = port.findTicket(parameters);
                            if (tk != null) {
                                rowNum = createSheetHeader(sheet, ++rowNum, tk.getReturn());
                            }
                        } catch (Exception ex) {
                            util.LogErr(ex.getMessage());
                        }
                        p.progress(i++);
                    }
                    // записываем созданный в памяти Excel документ в файл
                    try (FileOutputStream f = new FileOutputStream(saveFile)) {
                        workbook.write(f);
                    } catch (IOException ex) {
                        util.LogErr(ex.getMessage());
                    }
                    p.finish();
                    lines.clear();
                }
            };
            Thread t = new Thread(new Task(saveFile, lines));
            t.start();
        }
    }
}
