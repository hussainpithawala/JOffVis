package joffice.gui.listener;

import joffice.gui.frame.JOfficeFrame;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.Property;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.Section;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;

public class DocumentSummaryListener implements POIFSReaderListener {
  private JOfficeFrame jOfficeFrame;
  public DocumentSummaryListener(JOfficeFrame jOfficeFrame) {
    this.jOfficeFrame = jOfficeFrame;
  }
  public void processPOIFSReaderEvent(POIFSReaderEvent event) {
    StringBuilder documentSummary = new StringBuilder();
    DocumentSummaryInformation dsi = null;
    try {
      dsi = (DocumentSummaryInformation) PropertySetFactory.create(event
          .getStream());
    } catch (Exception ex) {
      throw new RuntimeException("Property set stream \"" + event.getPath()
          + event.getName() + "\": " + ex);
    }
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
