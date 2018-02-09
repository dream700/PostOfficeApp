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
import org.openide.util.Exceptions;
import ru.russianpost.siberia.GetTicket;
import ru.russianpost.siberia.GetTicketResponse;
import ru.russianpost.siberia.GetTicketSessionSRV;
import ru.russianpost.siberia.GetTicketSessionSRV_Service;
import ru.russianpost.siberia.Viewhistory;
import ru.russianpost.siberia.FindTicket;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketSRV;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketSRV_Service;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTickets;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetBatchTicketsResponse;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetReadyAnswer;
import ru.russianpost.siberia.maveneeticketlibrary.api.GetReadyAnswerResponse;
import ru.russianpost.siberia.Historyrecord;
import ru.russianpost.siberia.Ticket;
import ru.russianpost.siberia.ViewHistorySERV;
import ru.russianpost.siberia.ViewHistorySERV_Service;

/**
 *
 * @author andy
 */
@FunctionalInterface
interface Infotrace {

    void add(String data);
}

@FunctionalInterface
interface SetButtom {

    void SetButtom(boolean b);
}

public final class FactoryService {

    File reqfile;
    List<String> lines;
    boolean isStart;
    Infotrace funcInfotrace;
    SetButtom funcSetButtom;

    enum REQUEST {
        REQ_SERVER, REQ_DB, REQ_SERVER_DB
    }

    public FactoryService(Infotrace trace, SetButtom sbuttom) {
        lines = new ArrayList<>();
        funcInfotrace = trace;
        funcSetButtom = sbuttom;
        isStart = false;
    }

    private void addlog(String data) {
        funcInfotrace.add(data);
    }

    private void SetButtom(boolean b) {
        funcSetButtom.SetButtom(b);
        isStart = !b;
    }

    /*
    Читаем из файла данные и записываем в таблицу ticket
     */
    protected List<String> readFromFileTickets(List<String> lines, File file) {
        BufferedReader br = null;
        String line;
        String cvsSplitBy = "[|;]";
        try {
            br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
            while ((line = br.readLine()) != null) {
                if (!validateTicketFormat(lines, line)) {
                    String[] data = line.replaceAll("\"", "").split(cvsSplitBy);
                    for (String string : data) {
                        validateTicketFormat(lines, string);
                    }
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

    private boolean validateTicketFormat(List<String> lines, String barcode) {
        barcode = barcode.replace("\"", "");
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
//                            validateTicketFormat(tasklines, formatter.formatCellValue(cell));
//                            break;
//                        case STRING:
//                            validateTicketFormat(tasklines, formatter.formatCellValue(cell));
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

    private class WaittoAnswer extends SwingWorker<Void, String> {

        private final List<String> request;
        private final GetBatchTicketSRV_Service service;
        private final GetBatchTicketSRV port;
        private final GetReadyAnswer parameters;

        public WaittoAnswer(List<String> reqticket) {
            request = new ArrayList<>();
            request.addAll(reqticket);
            service = new GetBatchTicketSRV_Service();
            port = service.getGetBatchTicketSRVPort();
            parameters = new GetReadyAnswer();
        }

        @Override
        protected Void doInBackground() throws Exception {
            boolean f = true;
            while (f) {
                Thread.sleep(1000L); //10 sec
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
            addlog("Ok!\n");
            GetAnswerServer(lines);
            request.clear();
            SetButtom(true);
        }

        @Override
        protected void process(List<String> chunks) {
            for (String chunk : chunks) {
                addlog(chunk);
            }
        }

    }

    public void LoadDataOfServer(REQUEST req) {
        if (!isStart) {
            SetButtom(false);
            if ((reqfile = GetFileRequest(false, new FileNameExtensionFilter("TXT & CSV & XLS Files", "txt", "csv", "xls", "xlsx"))) != null) {
                util.changeCursorWaitStatus(true);
                try {
                    String filename = reqfile.getName().toLowerCase();
                    addlog("Загружаем фйал с ШПИ: " + filename + "\n");
                    util.LogDebug("File name - " + filename);
                    lines.clear();
                    if ((filename.indexOf("txt") > 0) | (filename.indexOf("csv") > 0)) {
                        addlog("Определен текстовый формат\n");
                        lines = readFromFileTickets(lines, reqfile);
                    } else if (filename.indexOf("xls") > 0) {
                        addlog("Определен excel формат\n");
                        lines = readFromExcelTickets(lines, reqfile);
                    } else {
                        addlog("Формат не определен\n");
                        addlog("Завершено\n");
                        SetButtom(true);
                    }
                    addlog("Загружаем из БД " + String.valueOf(lines.size()) + " ШПИ\n");
                    if (lines.size() > 0) {
                        switch (req) {
                            case REQ_SERVER:
                                List<String> request = GetBatchTicketService(lines);
                                addlog("Ожидаем ответа от сервера");
                                WaittoAnswer task = new WaittoAnswer(request);
                                task.execute();
                                break;
                            case REQ_DB:
                                GetAnswerServer(lines);
                                addlog("Готово!\n");
                                SetButtom(true);
                                break;
                            case REQ_SERVER_DB:
                                GetAnswerServerDetail(lines);
                                addlog("Готово!\n");
                                SetButtom(true);
                        }
                    } else {
                        addlog("Запрашивать не чего\n");
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
    }

    /*
    Отправляем данные на сервер
     */
    protected List<String> GetBatchTicketService(List<String> lines) {
        List<String> result = null;
        try { // Call Web Service Operation
            addlog("Отправляем запрос из " + String.valueOf(lines.size()) + " ШПИ на сервер...");
            GetBatchTicketSRV_Service service = new GetBatchTicketSRV_Service();
            GetBatchTicketSRV port = service.getGetBatchTicketSRVPort();
            GetBatchTickets parameters = new GetBatchTickets();
            parameters.getTickets().addAll(lines);
            GetBatchTicketsResponse resp = port.getBatchTickets(parameters);
            addlog("отправили\n");
            util.LogDebug("Get server done.");
            result = resp.getReturn();
        } catch (Exception ex) {
            addlog("ошибка\n");
            util.LogErr(ex.getMessage());
        }
        return result;
    }

    /*
    Получаем данные с сервера
     */
    class TaskDB extends SwingWorker<Void, Integer> {

        private final List<String> tasklines;
        private final String fout;

        TaskDB(String f, List<String> l) {
            fout = f;
            tasklines = new ArrayList<>();
            tasklines.addAll(l);
        }

        @Override
        protected Void doInBackground() throws Exception {
            ProgressHandle p = ProgressHandle.createHandle("Загрузка ШПИ с сервера");
            p.start(tasklines.size());
            // создание самого excel файла в памяти
            XSSFWorkbook workbook = new XSSFWorkbook();
            // создание листа с названием "Просто лист"
            XSSFSheet sheet = workbook.createSheet("ШПИ");
            // счетчик для строк
            int rowNum = 1;
            // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue("ШПИ");
            row.createCell(1).setCellValue("Финал");
            row.createCell(2).setCellValue("Индекс обработки");
            row.createCell(3).setCellValue("Индекс направления");
            row.createCell(4).setCellValue("Дата операции");
            row.createCell(5).setCellValue("ID Операции");
            row.createCell(6).setCellValue("Операция");
            row.createCell(7).setCellValue("ID Атрибута");
            row.createCell(8).setCellValue("Атрибут");
            row.createCell(9).setCellValue("Часы");

            ViewHistorySERV_Service service = new ViewHistorySERV_Service();
            ViewHistorySERV port = service.getViewHistorySERVPort();
            FindTicket parameters = new FindTicket();
            for (String line : tasklines) {
                try { // Call Web Service Operation
                    Ticket tk = port.findTicket(line);
                    if (tk != null) {
                        util.LogDebug(tk.getBarcode());
                        for (Historyrecord h : tk.getHistoryrecord()) {
                            Row r1 = sheet.createRow(++rowNum);
                            r1.createCell(0).setCellValue(tk.getBarcode());
                            r1.createCell(1).setCellValue(tk.isIsFinal());
                            r1.createCell(2).setCellValue(tk.getRecpIndex());
                            r1.createCell(3).setCellValue(h.getDestinationAddressIndex());
                            r1.createCell(4).setCellValue(h.getOperDate().toString());
                            r1.createCell(5).setCellValue(h.getOperTypeID());
                            r1.createCell(6).setCellValue(h.getOperTypeName());
                            r1.createCell(7).setCellValue(h.getOperAttrID());
                            r1.createCell(8).setCellValue(h.getOperAttrName());
                            r1.createCell(9).setCellValue(h.getOperatonDelta() / 60);
                        }
                        publish(rowNum);
                    }
                } catch (Exception ex) {
                    util.LogErr(ex.getMessage());
                }
                p.progress(rowNum);
            }
            p.finish();
            // записываем созданный в памяти Excel документ в файл
            File saveFile = new File(reqfile.getAbsolutePath()
                    .substring(0, reqfile.getAbsolutePath().lastIndexOf(
                            File.separator)) + File.separator + fout);
            util.LogDebug("Output file name" + saveFile.getName());
            try (FileOutputStream f = new FileOutputStream(saveFile)) {
                workbook.write(f);
                f.close();
            } catch (IOException ex) {
                util.LogErr(ex.getMessage());
            }
            return null;
        }

        @Override
        protected void done() {
            addlog("Ok!\n");
            util.LogDebug("Load final");
            tasklines.clear();
        }

        @Override
        protected void process(List<Integer> chunks) {
            for (Integer chunk : chunks) {
                addlog(".");
                util.LogDebug("row:" + chunk.toString() + "\n");
            }
        }

    }

    private void GetAnswerServer(List<String> lines) {
        String[] filename = reqfile.getName().split(Pattern.quote("."));
        if ("xlsx".equals(filename[1])) {
            filename[0] = filename[0] + "$";
        }
        String sfile = filename[0] + ".xlsx";
        addlog("Выгружаем в файл " + sfile + "\n");
        TaskDB task = new TaskDB(sfile, lines);
        task.execute();
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
        addlog("Выгружаем в файл " + saveFile.getName() + "\n");
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
                int rowNum = 1;
                // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
                Row row = sheet.createRow(rowNum);
                row.createCell(0).setCellValue("ШПИ");
                row.createCell(1).setCellValue("Финал");
                row.createCell(2).setCellValue("Индекс обработки");
                row.createCell(3).setCellValue("Индекс направления");
                row.createCell(4).setCellValue("Дата операции");
                row.createCell(5).setCellValue("Операция");
                row.createCell(6).setCellValue("Атрибут");
                row.createCell(7).setCellValue("Часы");
                GetTicketSessionSRV_Service service = new GetTicketSessionSRV_Service();
                GetTicketSessionSRV port = service.getGetTicketSessionSRVPort();
                GetTicket parameters = new GetTicket();
                for (String line : lines) {
                    try { // Call Web Service Operation
                        parameters.setBarcode(line);
                        GetTicketResponse tk = port.getTicket(parameters);
                        if (tk != null) {
                            for (Viewhistory h : tk.getReturn()) {
                                Row r = sheet.createRow(++rowNum);
                                r.createCell(0).setCellValue(h.getBarcode());
                                r.createCell(1).setCellValue(h.getOperationaddressIndex());
                                r.createCell(2).setCellValue(h.getDestinationaddressIndex());
                                r.createCell(3).setCellValue(h.getOperdate().toString());
                                r.createCell(4).setCellValue(h.getOpertypeid());
                                r.createCell(5).setCellValue(h.getNameattr());
                                r.createCell(6).setCellValue(h.getNametype());
                                r.createCell(7).setCellValue(h.getOperatondelta() / 60);
                            }
                        }
                    } catch (Exception ex) {
                        util.LogErr(ex.getMessage());
                    }
                    p.progress(rowNum);
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
