using System.Collections.Generic;

namespace Sample.DocumentGenerator
{
  public interface IDocumentGenerator
  {
    byte[] GenerateDocument(IDictionary<string, string> values, string fileName);
  }
}
