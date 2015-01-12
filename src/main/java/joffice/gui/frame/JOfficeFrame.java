package joffice.gui.frame;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingUtilities;
import javax.swing.border.EmptyBorder;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.filechooser.FileSystemView;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import joffice.gui.filter.ExtensionFileFilter;
import joffice.gui.model.RecordTreeNode;
import joffice.gui.model.RecordTreeTableModel;
import joffice.poi.listener.MyPOIFSReaderListener;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.eventfilesystem.POIFSReader;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.jdesktop.swingx.JXTreeTable;
import org.jdesktop.swingx.treetable.TreeTableModel;

@SuppressWarnings("serial")
public class JOfficeFrame extends JFrame {
  
  private static final String XLS = ".xls";
  private static final int DEFAULT_WIDTH = 800;
  private static final int DEFAULT_HEIGHT = 600;
  
  private String lastOpenedDirectory = null;
  private JXTreeTable treeTable = new JXTreeTable();
  private JTextArea descriptionArea = new JTextArea();
  
  /**
   * A description pane holds the description area
   */
  private JScrollPane descriptionPane = new JScrollPane(descriptionArea);
  /**
   * A tree table pane holds the treeTable for the records
   */
  private JScrollPane treeTablePane = new JScrollPane(treeTable);
  /**
   * A main pane which is split horizontally for holding two panes
   */
  private JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, treeTablePane, descriptionPane);
  
  /**
   * A content pane to hold the splitPane
   */
  private JPanel contentPanel = new JPanel();
  
  public JOfficeFrame() {
    this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    this.setVisible(true);
    setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
    setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    setBounds(100, 100, 450, 300);
    
    splitPane.setOneTouchExpandable(true);
      
    contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
    contentPanel.setLayout(new BorderLayout(0, 0));
    contentPanel.add(splitPane, BorderLayout.CENTER);
    
    splitPane.setResizeWeight(0.5f);
    
    addMenu();
    addTreeTable();
  }

  private void addTreeTable() {
    descriptionArea.setEditable(false); // set textArea non-editable
    descriptionPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
    
    treeTable.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
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
          descriptionArea.setText(record.toString());
        } else {
          descriptionArea.setText(null);
        }
      }
    });
   
    // A scroll pane for the table
    treeTablePane = new JScrollPane(treeTable);
    JSplitPane mainPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, treeTablePane, descriptionPane);
    
    mainPane.setDividerLocation(0.5);
    add(mainPane, BorderLayout.CENTER);
  }
  
  private void updateTreeTable(List<Record> records) {
    setTitle("Record Table");
    TreeTableModel treeTableModel = new RecordTreeTableModel(records);
    treeTable.setTreeTableModel(treeTableModel);
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
      // prompt the user for a xls file
      JFileChooser chooser = new JFileChooser();
      
      if (lastOpenedDirectory != null) {
        chooser.setCurrentDirectory(new File(lastOpenedDirectory));
      } else {
        chooser.setCurrentDirectory(FileSystemView.getFileSystemView().getDefaultDirectory());
      }
      ExtensionFileFilter filter = new ExtensionFileFilter();
      filter.addExtension(XLS);
      filter.setDescription("XLS Files");
      chooser.setFileFilter(filter);
      int r = chooser.showOpenDialog(JOfficeFrame.this);
      if (r == JFileChooser.APPROVE_OPTION) {
        File xlsFile = chooser.getSelectedFile();
        lastOpenedDirectory = chooser.getCurrentDirectory().getPath();
        String xlsName = xlsFile.getPath();
        loadXLSFile(xlsName);
      }
    }
  }
  /**
   * Parse the xls file and populate the table
   */
  private void loadXLSFile(String xlsFile) {
    try {
      System.out.println("SummaryInformation Default Stream Name " + SummaryInformation.DEFAULT_STREAM_NAME);
      POIFSReader r = new POIFSReader(); 
      r.registerListener(new MyPOIFSReaderListener(), SummaryInformation.DEFAULT_STREAM_NAME);
      r.read(new FileInputStream(xlsFile));
       
      NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(xlsFile));
      
      DirectoryNode directoryNode = fs.getRoot();
      Iterator<Entry> entries = directoryNode.getEntries();
      while (entries.hasNext()) {
        Entry entry = entries.next();
        System.out.println(entry.getName());
      }
      
      String workbookName = HSSFWorkbook
          .getWorkbookDirEntryName(directoryNode);
      
      InputStream stream = directoryNode
          .createDocumentInputStream(workbookName);
      List<Record> records = RecordFactory.createRecords(stream);
      updateTreeTable(records);
      fs.close();
    } catch (IOException e) {
      // TODO Auto-generated catch block
      JOptionPane.showMessageDialog(null, "Unable to open file");
      e.getMessage();
      e.printStackTrace();
    }
  }
}
