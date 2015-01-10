package joffice.gui.model;

import java.util.List;

import javax.swing.table.AbstractTableModel;

import org.apache.poi.hssf.record.Record;

@SuppressWarnings("serial")
public class RecordTableModel extends AbstractTableModel {

  public RecordTableModel(List<Record> records) {
    this.records = records;
  }
  
  public int getRowCount() {
    return records.size();
  }
  
  public int getColumnCount() {
    return 3;
  }
  
  public Object getValueAt(int r, int c) {
    Record record = records.get(r);
    Object value = null;
    switch(c) {
    case 0:
      value = record.getClass().getSimpleName();
      break;
    case 1:
      value = record.getSid();
      break;
    case 2:
      value = record.getRecordSize();
      break;
    }
    return value;
  }
  
  public String getColumnName(int c) {
    StringBuilder columnName = new StringBuilder();
    switch(c) {
    case 0:
      columnName.append("Name");
      break;
    case 1:
      columnName.append("Sid");
      break;
    case 2:
      columnName.append("Size");
      break;
    }
    return columnName.toString();
  }
  
  public String getRowDescription(int row) {
    return records.get(row).toString();
  }
  
  private List<Record> records;
}