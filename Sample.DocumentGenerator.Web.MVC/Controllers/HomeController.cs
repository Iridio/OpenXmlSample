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

    public FileContentResult GenerateDocument()
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
      IDocumentGenerator wordGen = new WordGenerator();
      var result = wordGen.GenerateDocument(values,  Server.MapPath("~/Content/Files/WordTest.docx"));
      return File(result, ("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), "WordTest.docx");
    }
  }
}
