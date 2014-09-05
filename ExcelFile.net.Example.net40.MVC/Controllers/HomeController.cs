using System.Web.Mvc;

namespace ExcelFile.net.Example.net40.MVC.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Export()
        {
            var excel = new ExcelFile();
            excel.Sheet("test").Row().Cell("111");
            excel.Save(Response, "测试.xls");
            return new EmptyResult();
        }
    }
}