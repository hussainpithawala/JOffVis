package joffice.strategy;

import java.io.IOException;
import java.util.Arrays;

import joffice.gui.frame.JOfficeFrame;

import org.apache.poi.hssf.OldExcelFormatException;

import com.google.common.base.Predicate;
import com.google.common.collect.Iterables;

public class Context {
  public enum FileExtensions {
    // Supported Extension
    XLS("xls"), DOC("doc"), PPT("ppt");
    /**
     * Returns the Context on the basis of value.
     * 
     * @param extension
     * @return Context.FileExtensions
     */
    public static Context.FileExtensions getContext(final String extension) {
      Context.FileExtensions fileExtension = Iterables.find(
          Arrays.asList(Context.FileExtensions.values()),
          new Predicate<Context.FileExtensions>() {
            public boolean apply(Context.FileExtensions fileExtension) {
              if (fileExtension.getName().equals(extension)) {
                return true;
              } else {
                return false;
              }
            }
          });
      return fileExtension;
    }

    private final String name;

    public String getName() {
      return this.name;
    }

    FileExtensions(String name) {
      this.name = name;
    }
  }

  private Strategy strategy;
  private String fileName;
  private JOfficeFrame jOfficeFrame;

  public Context(Strategy strategy, JOfficeFrame jOfficeFrame, String fileName) {
    this.strategy = strategy;
    this.jOfficeFrame = jOfficeFrame;
    this.fileName = fileName;
  }

  public void execute() {
    if (strategy != null) {
      try {
        strategy.execute(fileName, jOfficeFrame);
      } catch (OldExcelFormatException e) {
        e.printStackTrace();
        jOfficeFrame.showExceptionMessage(e);
      } catch (IOException e) {
        e.printStackTrace();
        jOfficeFrame.showExceptionMessage(e);
      }
    }
  }
}
