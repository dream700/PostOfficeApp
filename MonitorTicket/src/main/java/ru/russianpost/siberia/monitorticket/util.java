/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.russianpost.siberia.monitorticket;

import java.awt.Component;
import java.awt.Cursor;
import javax.swing.JFrame;
import org.openide.util.Mutex;
import org.openide.windows.IOProvider;
import org.openide.windows.InputOutput;
import org.openide.windows.WindowManager;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 * @author andy
 */
public class util {

    
    /*
    Дополнительные функции
     */
    public static String getTagValue(String sTag, Element eElement) {
        NodeList nlList = eElement.getElementsByTagName(sTag).item(0).getChildNodes();
        Node nValus = (Node) nlList.item(0);
        return nValus.getNodeValue();
    }

    public static String getValue(Element element) {
        String ret = "";
        if (element.hasChildNodes()) {
            ret = element.getChildNodes().item(0).getNodeValue();
        }
        return ret;
    }

    public static void changeCursorWaitStatus(final boolean isWaiting) {
        Mutex.EVENT.writeAccess(new Runnable() {
            public void run() {
                try {
                    JFrame mainFrame
                            = (JFrame) WindowManager.getDefault().getMainWindow();
                    Component glassPane = mainFrame.getGlassPane();
                    if (isWaiting) {
                        glassPane.setVisible(true);

                        glassPane.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
                    } else {
                        glassPane.setVisible(false);

                        glassPane.setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
                    }
                } catch (Exception e) {
                    // probably not worth handling 
                }
            }
        });
    }

    public static void LogDebug(String msg) {
        InputOutput io = IOProvider.getDefault().getIO("Проверка ШПИ", false);
        io.getOut().println(msg);
        io.getOut().close();
//        io.select();
    }

    public static void LogErr(String msg) {
        InputOutput io = IOProvider.getDefault().getIO("Проверка ШПИ", false);
        io.getErr().println(msg);
        io.getErr().close();
        io.select();
    }

}
