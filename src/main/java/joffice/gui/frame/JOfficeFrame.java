package joffice.gui.frame;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import joffice.gui.filter.ExtensionFileFilter;
import joffice.gui.model.RecordTreeNode;
import joffice.gui.model.RecordTreeTableModel;

import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.jdesktop.swingx.JXTreeTable;
import org.jdesktop.swingx.treetable.TreeTableModel;

@SuppressWarnings("serial")
public class JOfficeFrame extends JFrame {
  public JOfficeFrame() {
    this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    this.setVisible(true);
    setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
    addMenu();
  }

  private void addTreeTable(List<Record> records) {
    setTitle("Record Table");
    TreeTableModel treeTableModel = new RecordTreeTableModel(records);
    final JXTreeTable treeTable = new JXTreeTable(treeTableModel);

    treeTable.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
    
    final JTextArea recordDescription = new JTextArea();
    recordDescription.setEditable(false); // set textArea non-editable
    
    JScrollPane recordDescriptionScrollPane = new JScrollPane(recordDescription);
    recordDescriptionScrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);

    /**
     * Add a TreeSelectionListener which listens for the selected node  
     */
    treeTable.addTreeSelectionListener(new TreeSelectionListener() {
      public void valueChanged(TreeSelectionEvent e) {
        // TODO Auto-generated method stub
        TreePath path = treeTable.getTreeSelectionModel().getSelectionPath();
        if (path == null)
          return;
        RecordTreeNode recordTreeNode = (RecordTreeNode)path.getLastPathComponent();
        Record record = recordTreeNode.getRecord();
        if (record != null) {
          recordDescription.setText(record.toString());
        } else {
          recordDescription.setText(null);
        }
      }
    });
    
    // A scroll pane for the table
    JScrollPane scrollPane = new JScrollPane(treeTable);
    JSplitPane outerPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, scrollPane, recordDescriptionScrollPane);
    add(outerPane, BorderLayout.CENTER);
  }

  private void addMenu() {
    // add the menu and the Open and Exit Menu items
    JMenuBar menuBar = new JMenuBar();
    JMenu menu = new JMenu("File");

    JMenuItem openItem = new JMenuItem("Open");
    menu.add(openItem);
    openItem.addActionListener(new OpenAction());

    JMenuItem exitItem = new JMenuItem("Exit");
    menu.add(exitItem);

    exitItem.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent e) {
        System.exit(0);
      }
    });
    menuBar.add(menu);
    setJMenuBar(menuBar);
  }

  private class OpenAction implements ActionListener {
    public void actionPerformed(ActionEvent event) {
      // prompt the user for a zip file
      JFileChooser chooser = new JFileChooser();
      chooser.setCurrentDirectory(new File(
          "/Users/hussainp/Documents/Sheet-Issues"));
      ExtensionFileFilter filter = new ExtensionFileFilter();
      filter.addExtension(".xls");
      filter.setDescription("XLS Files");
      chooser.setFileFilter(filter);
      int r = chooser.showOpenDialog(JOfficeFrame.this);
      if (r == JFileChooser.APPROVE_OPTION) {
        String xlsName = chooser.getSelectedFile().getPath();
        System.out.println(xlsName);
        loadXLSFile(xlsName);
      }
    }
  }
  /**
   * Parse the xls file and populate the table
   */
  public void loadXLSFile(String xlsFile) {
    try {
      /*
       * POIFSReader r = new POIFSReader(); r.registerListener(new
       * MyPOIFSReaderListener(), "\005SummaryInformation"); r.read(new
       * FileInputStream(fileName));
       */
      NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(xlsFile));
      
      DirectoryNode directoryNode = fs.getRoot();

      String workbookName = HSSFWorkbook
          .getWorkbookDirEntryName(directoryNode);
      
      InputStream stream = directoryNode
          .createDocumentInputStream(workbookName);
      List<Record> records = RecordFactory.createRecords(stream);
      addTreeTable(records);
      fs.close();
    } catch (IOException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
  
  private static final int DEFAULT_WIDTH = 400;
  private static final int DEFAULT_HEIGHT = 200;
}
