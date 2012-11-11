using System.Web;
using System.Web.Mvc;

namespace Sample.DocumentGenerator.Web.MVC
{
  public class FilterConfig
  {
    public static void RegisterGlobalFilters(GlobalFilterCollection filters)
    {
      filters.Add(new HandleErrorAttribute());
    }
  }
}