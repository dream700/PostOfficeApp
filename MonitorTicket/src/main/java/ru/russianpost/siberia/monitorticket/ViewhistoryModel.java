package ru.russianpost.siberia.monitorticket;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import javax.swing.table.AbstractTableModel;
import ru.russianpost.siberia.maveneeticketlibrary.api.Viewhistory;

/**
 *
 * @author Andrey.Isakov
 */
public class ViewhistoryModel extends AbstractTableModel {

    private final List<Viewhistory> viewhostory;
    private final String[] columnNames = new String[]{
        "Индекс", "Дата", "Тип операции", "Операция", "Часов"
    };
    private final Class[] columnClass = new Class[]{
        String.class, Date.class, String.class, String.class, Integer.class
    };

    public ViewhistoryModel(List<Viewhistory> viewhostory) {
        this.viewhostory = viewhostory;
    }

    public Date getOperDate(String operdate) {
        DateFormat sf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSX");            
        Date parseDate;
        try {
            parseDate = sf.parse(operdate);
        } catch (ParseException ex) {
            parseDate = new Date();
        }
        return parseDate;
    }
    
    
    @Override
    public String getColumnName(int column) {
        return columnNames[column];
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        return columnClass[columnIndex];
    }

    @Override
    public int getRowCount() {
        return viewhostory.size();
    }

    @Override
    public int getColumnCount() {
        return columnNames.length;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Viewhistory row = viewhostory.get(rowIndex);
        switch (columnIndex) {
            case 0:
                return row.getOperationaddressIndex();
            case 1:
                return getOperDate(row.getOperdate().toString());
            case 2:
                return row.getNametype();
            case 3:
                return row.getNameattr();
            case 4:
                return row.getOperatondelta() / 60;
            default:
                break;
        }
        return null;
    }
}
