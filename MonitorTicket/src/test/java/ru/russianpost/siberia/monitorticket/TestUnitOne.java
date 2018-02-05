/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.russianpost.siberia.monitorticket;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;
import org.openide.util.Exceptions;

/**
 *
 * @author andy
 */
public class TestUnitOne {

    FactoryService fs;
    List<String> lines;
    File file;
    String[] data = {"63010217232797\n",
        "63010217269113\n",
        "63010217272281\n",
        "63010217289807\n",
        "63010217289814\n",
        "63010217289852\n",
        "63010217291305\n",
        "63010217290780\n",
        "63010217290797\n",
        "63010217290803\n"};

    public TestUnitOne() {
        fs = new FactoryService(new Infotrace() {
            @Override
            public void add(String data) {
                System.out.println(data);
            }
        }, new SetButtom() {
            @Override
            public void SetButtom(boolean b) {
                System.out.println("SetButtom: " + String.valueOf(b));
            }
        });
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
        lines = new ArrayList<>();
        file = CreateFileTXT();
    }

    @After
    public void tearDown() {
        lines.clear();
        if (file.exists()) {
            file.delete();
        }
    }

    // TODO add test methods here.
    // The methods must be annotated with annotation @Test. For example:
    //
    public File CreateFileTXT() {
        File f = new File("test.txt");
        BufferedWriter writer = null;
        try {
            writer = new BufferedWriter(new FileWriter(f));
            writer.write("63097717779182\n");
            writer.write("63097717779199\n");
            writer.write("63097717779205\n");
            writer.write("63097717779236\n");
            writer.write("63097717779212\n");
            writer.write("63097717779229\n");

        } catch (IOException e) {
        } finally {
            try {
                if (writer != null) {
                    writer.close();
                }
            } catch (IOException e) {
            }
        }
        return f;
    }

    private void CreateBigData() {
        lines.clear();
        for (int i = 0; i < 6000; i++) {
            for (String string : data) {
                lines.add(string);
            }
        }
        System.out.println("Array is " + String.valueOf(lines.size()));
    }

    @Test
    public void TestLoadTXT() {
        List<String> result = null;
        result = fs.readFromFileTickets(lines, file);
        assertTrue("Array is equal", result == lines);
        assertEquals("Read file", result.size(), 6);
    }

    @Test
    public void WriteXLSX() {
        CreateBigData();
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
        for (String line : lines) {
            Row r1 = sheet.createRow(++rowNum);
            r1.createCell(0).setCellValue(line);
            r1.createCell(1).setCellValue(new Date().toString());
            r1.createCell(2).setCellValue("FALSE");
            for (String h : data) {
                Row r2 = sheet.createRow(++rowNum);
                r2.createCell(1).setCellValue(h);
                r2.createCell(2).setCellValue(new Date().toString());
                r2.createCell(3).setCellValue("FALSE");
            }
        }
        // записываем созданный в памяти Excel документ в файл
        File saveFile = new File(file.getAbsolutePath()
                .substring(0, file.getAbsolutePath().lastIndexOf(
                        File.separator)) + File.separator + "test.xlsx");

        try {
            FileOutputStream f = new FileOutputStream(saveFile);
            workbook.write(f);
            f.close();
        } catch (IOException ex) {
            Exceptions.printStackTrace(ex);
        }
    }
}
