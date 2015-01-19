package joffice.gui.filter;

import java.io.File;
import java.util.ArrayList;

import javax.swing.filechooser.FileFilter;

public class ExtensionFileFilter extends FileFilter {
  /**
   * Adds an extnesion that this file filter recognizes.
   * 
   * @param extension
   * a file extension (such as ".xls" or ".doc")
   */
  public void addExtension(String extension) {
    if (!extension.startsWith("."))
      extension = "." + extension;
    extensions.add(extension.toLowerCase());
  }

  /**
   * Sets a description for the file set that this file filter recognizes.
   * 
   * @param aDescription
   *          a description for the file set
   */
  public void setDescription(String description) {
    this.description = description;
  }

  @Override
  public boolean accept(File f) {
    if (f.isDirectory())
      return true;

    String name = f.getName().toLowerCase();

    // check if the file name ends with any of the extensions
    for (String e : extensions) {
      if (name.endsWith(e)) {
        return true;
      }
    }
    return false;
  }

  @Override
  public String getDescription() {
    return description;
  }

  private String description = "";
  private ArrayList<String> extensions = new ArrayList<String>();
}
