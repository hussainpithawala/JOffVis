package joffice.gui.listener;

import joffice.gui.frame.JOfficeFrame;

import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;

public class SummaryListener implements POIFSReaderListener {
  private JOfficeFrame jOfficeFrame;

  public SummaryListener(JOfficeFrame jOfficeFrame) {
    this.jOfficeFrame = jOfficeFrame;
  }

  public void processPOIFSReaderEvent(POIFSReaderEvent event) {
    SummaryInformation si = null;
    try {
      si = (SummaryInformation) PropertySetFactory.create(event.getStream());
    } catch (Exception ex) {
      throw new RuntimeException("Property set stream \"" + event.getPath()
          + event.getName() + "\": " + ex);
    }
    StringBuilder summary = new StringBuilder();
    summary.append("Title " + si.getTitle() + "\n");
    summary.append("Last Saving application " + si.getApplicationName() + "\n");
    summary.append("OS Version " + si.getOSVersion() + "\n");
    summary.append("Rev " + si.getRevNumber() + "\n");
    summary.append("Original Author " + si.getAuthor() + "\n");
    summary.append("Last Author " + si.getLastAuthor() + "\n");
    summary.append("Last Saved Date " + si.getLastSaveDateTime() + "\n");
    summary.append("Creation Date " + si.getCreateDateTime() + "\n");
    jOfficeFrame.getSummaryArea().setText(summary.toString());
  }
}