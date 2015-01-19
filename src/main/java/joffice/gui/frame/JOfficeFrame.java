package joffice.gui.frame;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import joffice.gui.listener.OpenActionListener;
import joffice.gui.model.RecordTreeNode;
import joffice.gui.model.RecordTreeTableModel;

import org.apache.poi.hssf.record.Record;
import org.jdesktop.swingx.JXTreeTable;
import org.jdesktop.swingx.treetable.TreeTableModel;

@SuppressWarnings("serial")
public class JOfficeFrame extends JFrame {
  private static final int DEFAULT_WIDTH = 800;
  private static final int DEFAULT_HEIGHT = 600;

  private String lastOpenedDirectory = null;
  
  public String getLastOpenedDirectory() {
    return lastOpenedDirectory;
  }

  public void setLastOpenedDirectory(String lastOpenedDirectory) {
    this.lastOpenedDirectory = lastOpenedDirectory;
  }

  private JXTreeTable treeTable = new JXTreeTable();

  /**
   * A description pane holds the description area
   */
  private JTextArea descriptionArea = new JTextArea();
  private JScrollPane descriptionPane = new JScrollPane(descriptionArea);

  /**
   * A Summary pane holds the summary area
   */
  private JTextArea summaryArea = new JTextArea();
  public JTextArea getSummaryArea() {
    return summaryArea;
  }

  private JScrollPane summaryPane = new JScrollPane(summaryArea);
  
  /**
   * A Document summary pane holds the document summary area
   */
  private JTextArea documentSummaryArea = new JTextArea();
  
  public JTextArea getDocumentSummaryArea() {
    return documentSummaryArea;
  }

  private JScrollPane documentSummaryPane = new JScrollPane(documentSummaryArea);
  private JTabbedPane tabbedPane = new JTabbedPane(SwingConstants.TOP, JTabbedPane.SCROLL_TAB_LAYOUT);
  
  
  /**
   * Right Split Pane holds the summary information and descriptionPane
   */
  private JSplitPane rightSplitPane = 
      new JSplitPane(JSplitPane.VERTICAL_SPLIT,descriptionPane, tabbedPane);

  /**
   * A tree table pane holds the treeTable for the records
   */
  private JScrollPane treeTablePane = new JScrollPane(treeTable);
  /**
   * A main pane which is split horizontally for holding two panes
   */
  private JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
      treeTablePane, rightSplitPane);

  /**
   * A content pane to hold the splitPane
   */
  private JPanel contentPanel = new JPanel();

  public JOfficeFrame() {
    this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    this.setVisible(true);
    this.setTitle("JOffVis");
    
    setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
    setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    setBounds(100, 100, 450, 300);

    rightSplitPane.setOneTouchExpandable(true);
    rightSplitPane.setResizeWeight(0.7f);
    
    splitPane.setOneTouchExpandable(true);
    splitPane.setResizeWeight(0.4f);
    
    contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
    contentPanel.setLayout(new BorderLayout(0, 0));
    contentPanel.add(splitPane, BorderLayout.CENTER);

    add(contentPanel, BorderLayout.CENTER);

    addMenu();
    addTreeTable();
    addTabs();
  }

  private void addTabs() {
    tabbedPane.add("Summary ", summaryPane);
    tabbedPane.add("Document-Summary", documentSummaryPane);
  }
  
  private void addTreeTable() {
    descriptionArea.setEditable(false); // set textArea non-editable
    descriptionPane
        .setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);

    treeTable.getSelectionModel().setSelectionMode(
        TreeSelectionModel.SINGLE_TREE_SELECTION);
    /**
     * Add a TreeSelectionListener which listens for the selected node
     */
    treeTable.addTreeSelectionListener(new TreeSelectionListener() {
      public void valueChanged(TreeSelectionEvent e) {
        // TODO Auto-generated method stub
        TreePath path = treeTable.getTreeSelectionModel().getSelectionPath();
        if (path == null)
          return;
        RecordTreeNode recordTreeNode = (RecordTreeNode) path
            .getLastPathComponent();
        Record record = recordTreeNode.getRecord();
        if (record != null) {
          descriptionArea.setText(record.toString());
        } else {
          descriptionArea.setText(null);
        }
      }
    });
  }

  public void updateTreeTable(List<Record> records) {
    TreeTableModel treeTableModel = new RecordTreeTableModel(records);
    treeTable.setTreeTableModel(treeTableModel);
  }

  private void addMenu() {
    // add the menu and the Open and Exit Menu items
    JMenuBar menuBar = new JMenuBar();
    JMenu menu = new JMenu("File");

    JMenuItem openItem = new JMenuItem("Open");
    menu.add(openItem);
    openItem.addActionListener(new OpenActionListener(this));

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
  
  public void showExceptionMessage(Exception exception) {
    JTextArea jta = new JTextArea(exception.getMessage());
    JScrollPane jsp = new JScrollPane(jta){
        @Override
        public Dimension getPreferredSize() {
            return new Dimension(480, 320);
        }
    };
    JOptionPane.showMessageDialog(
        null, jsp, "Error", JOptionPane.ERROR_MESSAGE);
  }
}
