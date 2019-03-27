using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [AllowAnonymous]
        public ActionResult DownloadTemplate()
        {
            try
            {
                
                    string FilePath = Server.MapPath("~/754Template/");
                    DirectoryInfo dr = new DirectoryInfo(FilePath);
                    if (!dr.Exists)
                    {
                        dr.Create();
                    }
                    string FileName = "TemplateFileFrom754DataToUpdateAMT.xls";

                    byte[] fileBytes = System.IO.File.ReadAllBytes(FilePath + FileName);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, FileName);
                 
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View();

            }

        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}