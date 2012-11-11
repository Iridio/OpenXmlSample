using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Sample.DocumentGenerator
{
  public class WordGenerator : IDocumentGenerator
  {
    public byte[] GenerateDocument(IDictionary<string, string> values, string fileName)
    {
      if (values == null)
        throw new ArgumentException("Missing dictionary values.");
      if (!File.Exists(fileName))
        throw new ArgumentException("File \"" + fileName + "\" do not exists");
      var tempFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
      return CreateFile(values, fileName, tempFileName);
    }

    internal enum DocumentSection { Main, Header, Footer };

    internal static byte[] CreateFile(IDictionary<string, string> values, string fileName, string tempFileName)
    {
      File.Copy(fileName, tempFileName);
      if (!File.Exists(tempFileName))
        throw new ArgumentException("Unable to create file: " + tempFileName);

      using (var doc = WordprocessingDocument.Open(tempFileName, true))
      {
        if (doc.MainDocumentPart.HeaderParts != null)
          foreach (var header in doc.MainDocumentPart.HeaderParts)
            ProcessBookmarksPart(values, DocumentSection.Header, header);

        ProcessBookmarksPart(values, DocumentSection.Main, doc.MainDocumentPart);

        if (doc.MainDocumentPart.FooterParts != null)
          foreach (var footer in doc.MainDocumentPart.FooterParts)
            ProcessBookmarksPart(values, DocumentSection.Footer, footer);
      }
      byte[] result = null;
      if (File.Exists(tempFileName))
      {
        result = File.ReadAllBytes(tempFileName);
        File.Delete(tempFileName);
      }
      return result;
    }

    internal static void ProcessBookmarksPart(IDictionary<string, string> values, DocumentSection documentSection, object section)
    {
      IEnumerable<BookmarkStart> bookmarks = null;
      switch (documentSection)
      {
        case DocumentSection.Main:
          {
            bookmarks = ((MainDocumentPart)section).Document.Body.Descendants<BookmarkStart>();
            break;
          }
        case DocumentSection.Header:
          {
            bookmarks = ((HeaderPart)section).RootElement.Descendants<BookmarkStart>();
            break;
          }
        case DocumentSection.Footer:
          {
            bookmarks = ((FooterPart)section).RootElement.Descendants<BookmarkStart>();
            break;
          }
      }
      if (bookmarks == null) 
        return;
      foreach (var bmStart in bookmarks)
      {
        //If the bookmark name is not in our list. Just continue with the loop
        if (!values.ContainsKey(bmStart.Name)) 
          continue;
        BookmarkEnd bmEnd = null;
        switch (documentSection)
        {
          case DocumentSection.Main:
            {
              bmEnd = (((MainDocumentPart)section).Document.Body.Descendants<BookmarkEnd>().Where(b => b.Id == bmStart.Id.ToString())).FirstOrDefault();
              break;
            }
          case DocumentSection.Header:
            {
              bmEnd = (((HeaderPart)section).RootElement.Descendants<BookmarkEnd>().Where(b => b.Id == bmStart.Id.ToString())).FirstOrDefault();
              break;
            }
          case DocumentSection.Footer:
            {
              bmEnd =(((FooterPart)section).RootElement.Descendants<BookmarkEnd>().Where(b => b.Id == bmStart.Id.ToString())).FirstOrDefault();
              break;
            }
        }
        //If we did not find anything just continue in the loop
        if (bmEnd == null) 
          continue;
        var rProp = bmStart.Parent.Descendants<Run>().Where(rp => rp.RunProperties != null).Select(rp => rp.RunProperties).FirstOrDefault();
        if (bmStart.PreviousSibling<Run>() == null && bmEnd.ElementsAfter().Count(e => e.GetType() == typeof (Run)) == 0)
        {
          bmStart.Parent.RemoveAllChildren<Run>();
        }
        else
        {
          var list = bmStart.ElementsAfter().Where(r => r.IsBefore(bmEnd)).ToList();
          var trRun = list.Where(rp => rp.GetType() == typeof (Run) && ((Run) rp).RunProperties != null).Select(rp => ((Run) rp).RunProperties).FirstOrDefault();
          if (trRun != null)
            rProp = (RunProperties) trRun.Clone();
          for (var n = list.Count(); n > 0; n--)
            list[n-1].Remove();
        }
        var bmText = values[bmStart.Name];
        if (!string.IsNullOrEmpty(bmText) && bmText.Contains(Environment.NewLine))
        {
          var insertElement = bmStart.Parent.PreviousSibling();
          var rows = bmText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
          foreach (var row in rows)
          {
            var np = new Paragraph();
            var nRun = new Run();
            if (rProp != null)
              nRun.RunProperties = (RunProperties) rProp.Clone();
            nRun.AppendChild(new Text(row));
            np.AppendChild(nRun);
            if (insertElement.Parent != null)
              insertElement.InsertAfterSelf(np);
            else
              insertElement.Append(np);
            insertElement = np;
          }
        }
        else
        {
          var nRun = new Run();
          if (rProp != null)
            nRun.RunProperties = (RunProperties) rProp.Clone();
          nRun.Append(new Text(bmText));
          bmStart.InsertAfterSelf(nRun);
        }
      }
    }
  }
}
