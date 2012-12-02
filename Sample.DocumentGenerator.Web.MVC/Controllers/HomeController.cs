using System.Collections.Generic;
using System.Web.Mvc;

namespace Sample.DocumentGenerator.Web.MVC.Controllers
{
  public class HomeController : Controller
  {
    public ActionResult Index()
    {
      return View();
    }

    public FileContentResult GenerateWordDocument()
    {
      //It's a sample, so no DI just the bits to let it works
      var values = new Dictionary<string, string>
                     {
                       {"Title", "Foo title"},
                       {"Cell1", "123"},
                       {"Cell2", "asd"},
                       {"Header", "my header"},
                       {"Footer", "my footer"}
                     };
      IDocumentGenerator genereator = new WordGenerator();
      var result = genereator.GenerateDocument(values, Server.MapPath("~/Content/Files/WordTest.docx"));
      return File(result, ("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), "WordTest.docx");
    }

    public FileContentResult GenerateExcelDocument()
    {
      var values = new Dictionary<string, string>
                     {
                       {"ValueA", "90"},
                       {"ValueB", "300"},
                       {"Name", "Jhon Doe"}
                     };
      IDocumentGenerator generator = new ExcelGenerator();
      var result = generator.GenerateDocument(values, Server.MapPath("~/Content/Files/ExcelTest.xlsx"));
      return File(result, ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), "ExcelTest.xlsx");
    }
  }
}
