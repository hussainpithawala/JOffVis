package joffice.strategy;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.Property;
import org.apache.poi.hpsf.Section;
import org.apache.poi.hpsf.SummaryInformation;

import joffice.gui.frame.JOfficeFrame;

public abstract class Strategy {
  abstract void execute(String fileName, JOfficeFrame jOfficeFrame) 
      throws FileNotFoundException, IOException;
  void processSummaryInformation(SummaryInformation summaryInformation, JOfficeFrame jOfficeFrame) {
    StringBuilder summary = new StringBuilder();
    summary.append("Title " + summaryInformation.getTitle() + "\n");
    summary.append("Last Saving application " + summaryInformation.getApplicationName() + "\n");
    summary.append("OS Version " + summaryInformation.getOSVersion() + "\n");
    summary.append("Rev " + summaryInformation.getRevNumber() + "\n");
    summary.append("Original Author " + summaryInformation.getAuthor() + "\n");
    summary.append("Last Author " + summaryInformation.getLastAuthor() + "\n");
    summary.append("Last Saved Date " + summaryInformation.getLastSaveDateTime() + "\n");
    summary.append("Creation Date " + summaryInformation.getCreateDateTime() + "\n");
    jOfficeFrame.getSummaryArea().setText(summary.toString());
  }
  
  void processDocumentSummaryInformation(DocumentSummaryInformation dsi, JOfficeFrame jOfficeFrame) {
    StringBuilder documentSummary = new StringBuilder();
    for (Section section : dsi.getSections()) {
      for (Property property : section.getProperties()) {
        documentSummary.append("--------------------------" + "\n");
        documentSummary.append("ID: " + property.getID() + "\n");
        documentSummary.append("Value: " + property.getValue() + "\n");
        documentSummary.append("Type: " + property.getType() + "\n");
        documentSummary.append("--------------------------" + "\n");
      }
      documentSummary.append("\n");
    }
    // TODO Current Document Summary only prints properties
    // Need to fill this area with some relevant information
    // documentSummaryArea.setText(documentSummary.toString());
    jOfficeFrame.getDocumentSummaryArea().setText(documentSummary.toString());
  }
}
