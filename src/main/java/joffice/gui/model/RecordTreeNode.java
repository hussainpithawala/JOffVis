package joffice.gui.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFormatException;

public class RecordTreeNode {
  private Record record;
  private RecordTreeNode parent;
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
  public boolean addChild(RecordTreeNode childNode) {
    childNode.setParent(this);
    return children.add(childNode);
  }
  
  public RecordTreeNode getChildAtIndex(int index) {
    if (index >= 0 && index < children.size()) {
      return children.get(index);
    } else {
      return null;
    }
  }
  
  public int getSize() {
    int size = 0;
    if (children.isEmpty() && record != null) {
      try {
        size = record.getRecordSize();
      } catch(RecordFormatException rfe) {
        rfe.printStackTrace();
      }
    } else {
      for (RecordTreeNode childNode : children) {
        size += childNode.getSize();
      }      
    }
    return size;
  }
  
  public int indexOfChild(RecordTreeNode child) {
    return children.indexOf(child);
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
  public RecordTreeNode getParent() {
    return parent;
  }
  public void setParent(RecordTreeNode parent) {
    this.parent = parent;
  }
}
