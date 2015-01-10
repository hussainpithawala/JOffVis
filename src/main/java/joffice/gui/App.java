package joffice.gui;

import javax.swing.JFrame;
import javax.swing.SwingUtilities;

import joffice.gui.frame.JOfficeFrame;

public class App 
{
    public static void main( String[] args )
    {
        SwingUtilities.invokeLater(new Runnable() {
          public void run() {
              JOfficeFrame frame = new JOfficeFrame();
              frame.setVisible(true);
              frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
          }
      });
    }
}


