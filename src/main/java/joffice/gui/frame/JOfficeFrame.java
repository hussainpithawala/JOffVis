package joffice.gui.frame;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.border.MatteBorder;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableModel;

import joffice.gui.filter.ExtensionFileFilter;
import joffice.gui.model.RecordTableModel;

import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

@SuppressWarnings("serial")
public class JOfficeFrame extends JFrame {
  public JOfficeFrame() {
    this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    this.setVisible(true);
    setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
    addMenu();
  }

  private void addTable(List<Record> records) {
    setTitle("Record Table");
    TableModel model = new RecordTableModel(records);
    final JTable table = new JTable(model) {
      public Component prepareRenderer(TableCellRenderer renderer, int row,
          int column) {
        Component c = super.prepareRenderer(renderer, row, column);
        JComponent jc = (JComponent) c;
        // Color row based on a cell value
        // Alternate row color
        if (!isRowSelected(row))
          c.setBackground(row % 2 == 0 ? getBackground() : Color.LIGHT_GRAY);
        else
          jc.setBorder(new MatteBorder(1, 0, 1, 0, Color.RED));
        // Use bold font on selected row
        return c;
      }
    };

    final JTextArea recordDescription = new JTextArea();
    recordDescription.setEditable(false); // set textArea non-editable
    
    JScrollPane recordDescriptionScrollPane = new JScrollPane(recordDescription);
    recordDescriptionScrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);

    /**
     * Add a ListSelectionListener for a row
     */
    ListSelectionListener listSelectionListener = new ListSelectionListener() {
      public void valueChanged(ListSelectionEvent event) {
        if (table.getSelectedRow() > -1) {
          RecordTableModel recordTableModel = (RecordTableModel)table.getModel();
          recordDescription.setText(recordTableModel.getRowDescription(table.getSelectedRow()));
        }
      }
    };
    
    table.getSelectionModel().addListSelectionListener(listSelectionListener);

    // A scroll pane for the table
    JScrollPane scrollPane = new JScrollPane(table);
    JSplitPane outerPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, scrollPane, recordDescriptionScrollPane);
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
        xlsName = chooser.getSelectedFile().getPath();
        loadXLSFile();
      }
    }

    /**
     * Scans the xls file and populates the table
     */

    public void loadXLSFile() {
      try {
        /*
         * POIFSReader r = new POIFSReader(); r.registerListener(new
         * MyPOIFSReaderListener(), "\005SummaryInformation"); r.read(new
         * FileInputStream(fileName));
         */
        NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(xlsName));
        DirectoryNode directoryNode = fs.getRoot();

        String workbookName = HSSFWorkbook
            .getWorkbookDirEntryName(directoryNode);

        InputStream stream = directoryNode
            .createDocumentInputStream(workbookName);

        List<Record> records = RecordFactory.createRecords(stream);
        addTable(records);
        fs.close();
      } catch (IOException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
      }
    }
  }

  private String xlsName;
  private static final int DEFAULT_WIDTH = 400;
  private static final int DEFAULT_HEIGHT = 200;
}
