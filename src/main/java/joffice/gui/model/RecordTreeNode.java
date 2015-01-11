package joffice.gui.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.record.Record;

public class RecordTreeNode {
  private Record record;
  private List<RecordTreeNode> children = new ArrayList<RecordTreeNode>();
  
  public RecordTreeNode() {
    
  }
  public RecordTreeNode(Record record){
    this.record = record;
  }
  public Record getRecord() {
    return record;
  }
  public void setRecord(Record record) {
    this.record = record;
  }
  public List<RecordTreeNode> getChildren() {
    return children;
  }
  public boolean addChildren(RecordTreeNode childNode) {
    return children.add(childNode);
  }
  
  public int getSize() {
    int size = 0;
    if (children.isEmpty() && record != null) {
      size = record.getRecordSize();
    } else {
      for (RecordTreeNode childNode : children) {
        size += childNode.getSize();
      }      
    }
    return size;
  }
  
  public int getCount() {
    int count = 0;
    if (children.isEmpty() && record != null) {
      count = 1;
    } else {
      for (RecordTreeNode childNode : children) {
        count += childNode.getCount();
      }      
    }
    return count;
  }
}
