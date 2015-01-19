package joffice.strategy;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import joffice.gui.frame.JOfficeFrame;

import org.apache.poi.hssf.record.RecordFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

class XLSStrategy extends Strategy {
  public void execute(String fileName, JOfficeFrame jOfficeFrame) 
      throws FileNotFoundException, IOException {
    
    NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(fileName));

    DirectoryNode directoryNode = fs.getRoot();
    
    String workbookName = HSSFWorkbook.getWorkbookDirEntryName(directoryNode);
    
    HSSFWorkbook workbook = new HSSFWorkbook(directoryNode, true);
    
    workbook.getDocumentSummaryInformation();
    workbook.getSummaryInformation();
    
    InputStream stream = directoryNode
        .createDocumentInputStream(workbookName);

    jOfficeFrame.updateTreeTable(RecordFactory.createRecords(stream));
    
    /**
     * Process Summary Information and Document SummaryInformation
     */
    processDocumentSummaryInformation(workbook.getDocumentSummaryInformation(), jOfficeFrame);
    processSummaryInformation(workbook.getSummaryInformation(), jOfficeFrame);
    fs.close();
  }
}
