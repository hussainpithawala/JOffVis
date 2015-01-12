package joffice.gui.model;

import java.util.List;

import org.apache.poi.hssf.record.Record;
import org.jdesktop.swingx.treetable.AbstractTreeTableModel;

public class RecordTreeTableModel extends AbstractTreeTableModel {
  private RecordTreeNode rootNode;
  private List<Record> rawRecords;

  public RecordTreeTableModel(List<Record> rawRecords) {
    this.rawRecords = rawRecords;
    rootNode = new RecordTreeNode();
    for (int index = 0; index < rawRecords.size(); ++index) {
      Record record = rawRecords.get(index);
      RecordTreeNode childNode = new RecordTreeNode(record);
      if (record.getClass().getSimpleName().equals("BOFRecord")) {
        int lastIndex = processChildTree(rootNode, childNode, index);
        index = lastIndex;
      } else {
        rootNode.addChildren(childNode);
      }
    }
  }

  /**
   * Process the complete child tree
   * @param the current RecordTreeNode, which basically refers to the BOFRecord
   * @index the running index of the rawRecords
   */
  private int processChildTree(RecordTreeNode parentNode, RecordTreeNode childNode, int index) {
    // Add a Dummy BIFF-RECORDS RecordTreeNode
    RecordTreeNode containerNode = new RecordTreeNode();
    
    // Add the BOF Record to the container Node
    containerNode.addChildren(childNode);
    
    int lastIndex = addChildrenToParent(containerNode, index + 1);
    index = lastIndex;
    
    // Add the EOF Record to the container Node
    childNode = new RecordTreeNode(rawRecords.get(index));
    containerNode.addChildren(childNode);
    
    // Add the dummy container node to the children node
    parentNode.addChildren(containerNode);
    
    return lastIndex;
  }
  /**
   * On encountering a BOF record this function needs to be called with a
   * RecordTreeNode, which holds the current BOF Record and index of the BOF
   * record from which we need to move forward till we either reach the end of
   * Raw Records list or EOF.
   * 
   * @param parentNode
   * @param prevIndex
   */
  private int addChildrenToParent(RecordTreeNode parentNode, int prevIndex) {
    int lastIndex = prevIndex;

    INNER: for (int index = prevIndex; index < rawRecords.size(); ++index) {
      Record record = rawRecords.get(index);
      RecordTreeNode childNode = new RecordTreeNode(record);

      if (record.getClass().getSimpleName().equals("EOFRecord")) {
        lastIndex = index;
        break INNER;
      } else if (record.getClass().getSimpleName().equals("BOFRecord")) {
        lastIndex = processChildTree(parentNode, childNode, index);
        index = lastIndex;
      }
      parentNode.addChildren(childNode);
    }
    return lastIndex;
  }

  public int getColumnCount() {
    return 3;
  }

  public Object getValueAt(Object node, int columnIndex) {
    RecordTreeNode recordTreeNode = ((RecordTreeNode) node);
    Record record = recordTreeNode.getRecord();
    Object value = null;
    if (record != null) {
      switch (columnIndex) {
      case 0:
        value = getRecordName(record);
        break;
      case 1:
        value = record.getSid();
        break;
      case 2:
        value = record.getRecordSize();
        break;
      }
    } else {
      switch (columnIndex) {
      case 0:
        value = "BIFF Records [" + recordTreeNode.getCount() + "]";
        break;
      case 1:
        value = "Container ";
        break;
      case 2:
        int size = recordTreeNode.getSize();
        value = size;
        break;
      }
    }
    return value;
  }

  public Object getChild(Object parent, int index) {
    RecordTreeNode parentRecordTreeNode = (RecordTreeNode) parent;
    return parentRecordTreeNode.getChildren().get(index);
  }

  public int getChildCount(Object parent) {
    RecordTreeNode parentRecordTreeNode = (RecordTreeNode) parent;
    return parentRecordTreeNode.getChildren().size();
  }

  public int getIndexOfChild(Object parent, Object child) {
    RecordTreeNode parentRecordTreeNode = (RecordTreeNode) parent;
    return parentRecordTreeNode.getChildren().indexOf((RecordTreeNode) child);
  }

  public String getColumnName(int columnIndex) {
    StringBuilder columnName = new StringBuilder();
    switch (columnIndex) {
    case 0:
      columnName.append("Name");
      break;
    case 1:
      columnName.append("Type");
      break;
    case 2:
      columnName.append("Size");
      break;
    }
    return columnName.toString();
  }
  
  public String getRecordName(Record record) {
    String value = null;
    String simpleName = record.getClass().getSimpleName();
    String baseName = simpleName.substring(0, simpleName.length() - 6);
    
    if (baseName.equals("Unknown")) {
      String description = record.toString();
      value = description.substring(1, description.indexOf(']'));
    } else {
      value = baseName;
    }
    return value;
  }
  
  @Override
  public Object getRoot()
  {
    return rootNode;
  }
}
