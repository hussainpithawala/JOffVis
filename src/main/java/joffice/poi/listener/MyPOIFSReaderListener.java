package joffice.poi.listener;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;

public class MyPOIFSReaderListener implements POIFSReaderListener
{
    public void processPOIFSReaderEvent(POIFSReaderEvent event)
    {
        SummaryInformation si = null;
        DocumentSummaryInformation dsi = null;
        try
        {
            si = (SummaryInformation)
                 PropertySetFactory.create(event.getStream());
            /*
            dsi = (DocumentSummaryInformation)
                PropertySetFactory.create(event.getStream());
           */     
        }
        catch (Exception ex)
        {
            throw new RuntimeException
                ("Property set stream \"" +
                 event.getPath() + event.getName() + "\": " + ex);
        }
        // Summary Information
        System.out.println("Title " + si.getTitle());
        System.out.println("Last Saving application " + si.getApplicationName());
        System.out.println("OS Version " + si.getOSVersion());
        System.out.println("Rev " + si.getRevNumber());
        System.out.println("Original Author " + si.getAuthor());
        System.out.println("Last Author " + si.getLastAuthor());
        System.out.println("Last Saved Date " + si.getLastSaveDateTime());
        System.out.println("Creation Date " + si.getCreateDateTime());
    }
}    