package joffice.strategy;


public class StrategyFactory {
  /**
   * Prepares the strategy to be executed by the context on the basis 
   * of the file extension
   * @param extension
   * @return Strategy
   */
  public static Strategy prepare(final String extension) {
    Strategy strategy = null;
    switch(Context.FileExtensions.getContext(extension)) {
    case XLS :
      strategy = new XLSStrategy();
     break;
    case DOC:
      strategy = new DOCStrategy();
      break;
    case PPT:
      strategy = new PPTStrategy();
      break;
    default:
      break;
    }
    return strategy;
  }
}
