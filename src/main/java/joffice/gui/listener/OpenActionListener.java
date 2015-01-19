package joffice.gui.listener;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

import joffice.gui.filter.ExtensionFileFilter;
import joffice.gui.frame.JOfficeFrame;
import joffice.strategy.Context;
import joffice.strategy.Strategy;
import joffice.strategy.StrategyFactory;

import com.google.common.io.Files;

public class OpenActionListener implements ActionListener {
  private JOfficeFrame jOfficeFrame;

  public OpenActionListener(JOfficeFrame jOfficeFrame) {
    this.jOfficeFrame = jOfficeFrame;
  }

  public void actionPerformed(ActionEvent event) {
    JFileChooser chooser = new JFileChooser();

    if (jOfficeFrame.getLastOpenedDirectory() != null) {
      chooser.setCurrentDirectory(new File(jOfficeFrame
          .getLastOpenedDirectory()));
    } else {
      chooser.setCurrentDirectory(FileSystemView.getFileSystemView()
          .getDefaultDirectory());
    }

    ExtensionFileFilter extFileFilter = new ExtensionFileFilter();
    
    extFileFilter.addExtension("."+Context.FileExtensions.DOC.getName());
    extFileFilter.addExtension("."+Context.FileExtensions.XLS.getName());
    extFileFilter.addExtension("."+Context.FileExtensions.PPT.getName());
    extFileFilter.setDescription("DOC, XLS or PPT files");

    chooser.setFileFilter(extFileFilter);

    int r = chooser.showOpenDialog(jOfficeFrame);

    if (r == JFileChooser.APPROVE_OPTION) {
      File file = chooser.getSelectedFile();
      String extension = Files.getFileExtension(file.getName());
      jOfficeFrame.setLastOpenedDirectory(chooser.getCurrentDirectory()
          .getPath());
      String fileName = file.getPath();

      // Prepares a strategy to execute on the basis of selected file's extension
      Strategy strategy = StrategyFactory.prepare(extension);
      Context context = new Context(strategy, jOfficeFrame, fileName);
      context.execute();
    }
  }
}