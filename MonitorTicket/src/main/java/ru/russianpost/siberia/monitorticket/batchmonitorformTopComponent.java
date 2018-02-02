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
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;
import javax.swing.JFileChooser;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.netbeans.api.progress.ProgressHandle;
import org.netbeans.api.settings.ConvertAsProperties;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.util.Exceptions;
import org.openide.windows.TopComponent;
import org.openide.util.NbBundle.Messages;
import ru.russianpost.siberia.GetTicket;
import ru.russianpost.siberia.GetTicketResponse;
import ru.russianpost.siberia.GetTicketSessionSRV;
import ru.russianpost.siberia.GetTicketSessionSRV_Service;
import ru.russianpost.siberia.Viewhistory;
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

    private int createSheetHeader(XSSFSheet sheet, int rowNum, Ticket t) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(t.getBarcode());
        row.createCell(1).setCellValue(t.getDateFetch().toString());
        row.createCell(2).setCellValue(t.isIsFinal());
        for (Historyrecord h : t.getHistoryrecord()) {
            Row r = sheet.createRow(++rowNum);
            r.createCell(1).setCellValue(h.getOperationAddressIndex());
            r.createCell(2).setCellValue(h.getDestinationAddressIndex());
            r.createCell(3).setCellValue(h.getOperDate().toString());
            r.createCell(4).setCellValue(h.getOperTypeID());
            r.createCell(5).setCellValue(h.getOperTypeName());
            r.createCell(6).setCellValue(h.getOperAttrID());
            r.createCell(7).setCellValue(h.getOperAttrName());
            r.createCell(8).setCellValue(h.getOperatonDelta());
        }
        return rowNum;
    }

    private int createSheetHeaderDetail(XSSFSheet sheet, int rowNum, List<Viewhistory> t) {
        for (Viewhistory h : t) {
            Row r = sheet.createRow(++rowNum);
            r.createCell(0).setCellValue(h.getBarcode());
            r.createCell(1).setCellValue(h.getOperationaddressIndex());
            r.createCell(2).setCellValue(h.getDestinationaddressIndex());
            r.createCell(3).setCellValue(h.getOperdate().toString());
            r.createCell(4).setCellValue(h.getOpertypeid());
            r.createCell(5).setCellValue(h.getNameattr());
            r.createCell(6).setCellValue(h.getNametype());
            r.createCell(8).setCellValue(h.getOperatondelta());
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

        btLocal = new javax.swing.JButton();
        btFileLoad = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        InfoTrace = new javax.swing.JTextArea();
        btDetail = new javax.swing.JButton();

        setNextFocusableComponent(btFileLoad);

        org.openide.awt.Mnemonics.setLocalizedText(btLocal, org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btLocal.text")); // NOI18N
        btLocal.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btLocalActionPerformed(evt);
            }
        });

        org.openide.awt.Mnemonics.setLocalizedText(btFileLoad, org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btFileLoad.text")); // NOI18N
        btFileLoad.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btFileLoadActionPerformed(evt);
            }
        });

        InfoTrace.setColumns(20);
        InfoTrace.setRows(5);
        jScrollPane2.setViewportView(InfoTrace);

        org.openide.awt.Mnemonics.setLocalizedText(btDetail, org.openide.util.NbBundle.getMessage(batchmonitorformTopComponent.class, "batchmonitorformTopComponent.btDetail.text")); // NOI18N
        btDetail.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btDetailActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(btDetail, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btLocal, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btFileLoad, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 399, Short.MAX_VALUE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(16, 16, 16)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 143, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(btFileLoad)
                        .addGap(16, 16, 16)
                        .addComponent(btLocal)
                        .addGap(18, 18, 18)
                        .addComponent(btDetail)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
    }// </editor-fold>//GEN-END:initComponents

    private void SetButtom(boolean b) {
        btFileLoad.setEnabled(b);
        btLocal.setEnabled(b);
        btDetail.setEnabled(b);
    }

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
            InfoTrace.append("Ok!\n");
            GetAnswerServer(lines);
            lines.clear();
            btFileLoad.setEnabled(true);
            btLocal.setEnabled(true);
            btDetail.setEnabled(true);
        }

        @Override
        protected void process(List<String> chunks) {
            for (String chunk : chunks) {
                InfoTrace.append(chunk);
            }
        }

    }

    private void btFileLoadActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btFileLoadActionPerformed
        SetButtom(false);
        if ((reqfile = GetFileRequest(false, new FileNameExtensionFilter("TXT & CSV & XLS Files", "txt", "csv", "xls", "xlsx"))) != null) {
            util.changeCursorWaitStatus(true);
            try {
                String filename = reqfile.getName().toLowerCase();
                InfoTrace.append("Загружаем фйал с ШПИ: " + filename + "\n");
                util.LogDebug("File name - " + filename);
                lines.clear();
                if ((filename.indexOf("txt") > 0) | (filename.indexOf("csv") > 0)) {
                    InfoTrace.append("Определен текстовый формат\n");
                    lines = readFromFileTickets(lines, reqfile);
                } else if (filename.indexOf("xls") > 0) {
                    InfoTrace.append("Определен excel формат\n");
                    lines = readFromExcelTickets(lines, reqfile);
                } else {
                    InfoTrace.append("Формат не определен\n");
                    InfoTrace.append("Завершено\n");
                    SetButtom(true);
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
            util.changeCursorWaitStatus(false);
        } else {
            SetButtom(true);
        }
    }//GEN-LAST:event_btFileLoadActionPerformed

    private void btLocalActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btLocalActionPerformed
        if (btFileLoad.isEnabled()) {
            SetButtom(false);
            if ((reqfile = GetFileRequest(false, new FileNameExtensionFilter("TXT & CSV & XLS Files", "txt", "csv", "xls", "xlsx"))) != null) {
                util.changeCursorWaitStatus(true);
                try {
                    String filename = reqfile.getName().toLowerCase();
                    InfoTrace.append("Загружаем фйал с ШПИ: " + filename + "\n");
                    util.LogDebug("File name - " + filename);
                    lines.clear();
                    if ((filename.indexOf("txt") > 0) | (filename.indexOf("csv") > 0)) {
                        InfoTrace.append("Определен текстовый формат\n");
                        lines = readFromFileTickets(lines, reqfile);
                    } else if (filename.indexOf("xls") > 0) {
                        InfoTrace.append("Определен excel формат\n");
                        lines = readFromExcelTickets(lines, reqfile);
                    } else {
                        InfoTrace.append("Формат не определен\n");
                        InfoTrace.append("Завершено\n");
                        SetButtom(true);
                    }
                    InfoTrace.append("Загружаем из БД " + String.valueOf(lines.size()) + " ШПИ\n");
                    if (lines.size() > 0) {
                        GetAnswerServer(lines);
                        InfoTrace.append("Готово!\n");
                        SetButtom(true);
                    }
                } catch (Exception ex) {
                    Exceptions.printStackTrace(ex);
                }
                util.changeCursorWaitStatus(false);
            } else {
                SetButtom(true);
            }
        }
    }//GEN-LAST:event_btLocalActionPerformed

    private void btDetailActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btDetailActionPerformed
        if (btFileLoad.isEnabled()) {
            SetButtom(false);
            if ((reqfile = GetFileRequest(false, new FileNameExtensionFilter("TXT & CSV & XLS Files", "txt", "csv", "xls", "xlsx"))) != null) {
                util.changeCursorWaitStatus(true);
                try {
                    String filename = reqfile.getName().toLowerCase();
                    InfoTrace.append("Загружаем фйал с ШПИ: " + filename + "\n");
                    util.LogDebug("File name - " + filename);
                    lines.clear();
                    if ((filename.indexOf("txt") > 0) | (filename.indexOf("csv") > 0)) {
                        InfoTrace.append("Определен текстовый формат\n");
                        lines = readFromFileTickets(lines, reqfile);
                    } else if (filename.indexOf("xls") > 0) {
                        InfoTrace.append("Определен excel формат\n");
                        lines = readFromExcelTickets(lines, reqfile);
                    } else {
                        InfoTrace.append("Формат не определен\n");
                        InfoTrace.append("Завершено\n");
                        SetButtom(true);
                    }
                    InfoTrace.append("Загружаем из БД " + String.valueOf(lines.size()) + " ШПИ\n");
                    if (lines.size() > 0) {
                        GetAnswerServerDetail(lines);
                        InfoTrace.append("Готово!\n");
                        SetButtom(true);
                    }
                } catch (Exception ex) {
                    Exceptions.printStackTrace(ex);
                }
                util.changeCursorWaitStatus(false);
            } else {
                SetButtom(true);
            }
        }
    }//GEN-LAST:event_btDetailActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextArea InfoTrace;
    private javax.swing.JButton btDetail;
    private javax.swing.JButton btFileLoad;
    private javax.swing.JButton btLocal;
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
        String[] filename = reqfile.getName().split(Pattern.quote("."));
        if ("xlsx".equals(filename[1])) {
            filename[0] = filename[0] + "$";
        }
        String sfile = filename[0] + ".xlsx";
        File saveFile = new File(reqfile.getAbsolutePath()
                .substring(0, reqfile.getAbsolutePath().lastIndexOf(
                        File.separator)) + File.separator + sfile);
        InfoTrace.append("Выгружаем в файл " + saveFile.getName() + "\n");
        class Task implements Runnable {

            List<String> lines;
            File saveFile;

            Task(File f, List<String> l) {
                saveFile = f;
                lines = new ArrayList<>();
                lines.addAll(l);
            }

            @Override
            public void run() {
                ProgressHandle p = ProgressHandle.createHandle("Загрузка ШПИ с сервера");
                p.start(lines.size());
                // создание самого excel файла в памяти
                XSSFWorkbook workbook = new XSSFWorkbook();
                // создание листа с названием "Просто лист"
                XSSFSheet sheet = workbook.createSheet("ШПИ");
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
        }
        Thread t = new Thread(new Task(saveFile, lines));
        t.start();
    }

    /*
    Получаем данные с сервера
     */
    private void GetAnswerServerDetail(List<String> lines) {
        String[] filename = reqfile.getName().split(Pattern.quote("."));
        if ("xlsx".equals(filename[1])) {
            filename[0] = filename[0] + "$";
        }
        String sfile = filename[0] + ".xlsx";
        File saveFile = new File(reqfile.getAbsolutePath()
                .substring(0, reqfile.getAbsolutePath().lastIndexOf(
                        File.separator)) + File.separator + sfile);
        InfoTrace.append("Выгружаем в файл " + saveFile.getName() + "\n");
        class Task implements Runnable {

            List<String> lines;
            File saveFile;

            Task(File f, List<String> l) {
                saveFile = f;
                lines = new ArrayList<>();
                lines.addAll(l);
            }

            @Override
            public void run() {
                ProgressHandle p = ProgressHandle.createHandle("Загрузка ШПИ с сервера");
                p.start(lines.size());
                // создание самого excel файла в памяти
                XSSFWorkbook workbook = new XSSFWorkbook();
                // создание листа с названием "Просто лист"
                XSSFSheet sheet = workbook.createSheet("ШПИ");
                // счетчик для строк
                int rowNum = 0;
                // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
                Row row = sheet.createRow(rowNum);
                row.createCell(0).setCellValue("ШПИ");
                row.createCell(1).setCellValue("Дата");
                row.createCell(2).setCellValue("Финал");
                GetTicketSessionSRV_Service service = new GetTicketSessionSRV_Service();
                GetTicketSessionSRV port = service.getGetTicketSessionSRVPort();
                GetTicket parameters = new GetTicket();
                int i = 1;
                for (String line : lines) {
                    try { // Call Web Service Operation
                        parameters.setBarcode(line);
                        GetTicketResponse tk = port.getTicket(parameters);
                        if (tk != null) {
                            rowNum = createSheetHeaderDetail(sheet, ++rowNum, tk.getReturn());
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
