using System;
using System.IO;
using Microsoft.Office.Interop.Word;

class Program
{
  static void Main(string[] args)
  {
    object missing = Type.Missing;
    object saveChanges = false;

    _Application app = new ApplicationClass();

    WordDocumentCount total = new WordDocumentCount();
    foreach (string pattern in args)
    {
      string dir = Path.GetDirectoryName(pattern);
      string x = Path.GetFileName(pattern);
      if (dir.Length == 0)
        dir = ".";
      string[] files = Directory.GetFiles(dir, x);
      foreach (string file in files)
      {
        if (Path.GetExtension(file) != ".doc")
          continue;
        string absPath = Path.GetFullPath(file);
        WordDocumentCount results = CountFile(app, absPath);
        Console.WriteLine("{0} {1}", results, file);
        total += results;
      }
    }
    Console.WriteLine("{0} {1}", total, "total");

    app.Quit(ref saveChanges, ref missing, ref missing);
  }

  private static WordDocumentCount CountFile(_Application app, string filename)
  {
    object filename_ = filename;
    object readOnly = true;
    object m = Type.Missing;
    object saveChanges = false;
    object incFoot = false;
    object noEncodingDialog = true;

    int pages = 0;
    int words = 0;
    int characters = 0;
    try
    {
      _Document doc;
      doc = app.Documents.Open(ref filename_, ref m, ref readOnly, ref m,
        ref m, ref m, ref m, ref m, ref m, ref m, ref m, ref m, ref m,
        ref m, ref noEncodingDialog, ref m);
      pages = doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref incFoot);
      words = doc.ComputeStatistics(WdStatistic.wdStatisticWords, ref incFoot);
      characters = doc.ComputeStatistics(WdStatistic.wdStatisticCharacters,
        ref incFoot);
      doc.Close(ref saveChanges, ref m, ref m);
    }
    catch (Exception)
    {
    }

    return new WordDocumentCount(pages, words, characters);
  }
}

public class WordDocumentCount
{
  public int Pages;
  public int Words;
  public int Characters;

  public WordDocumentCount()
  {
  }

  public WordDocumentCount(int pages, int words, int characters)
  {
    this.Pages = pages;
    this.Words = words;
    this.Characters = characters;
  }

  public static WordDocumentCount operator +(WordDocumentCount x,
    WordDocumentCount y)
  {
    return new WordDocumentCount(x.Pages + y.Pages, x.Words + y.Words,
      x.Characters + y.Characters);
  }

  public override string ToString()
  {
    return string.Format("{0,7} {1,7} {2,7}", this.Pages, this.Words,
      this.Characters);
  }
}