using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace OfficeTestSpire.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            Document document = new Document();
            Paragraph paragraph = document.AddSection().AddParagraph();
            paragraph.AppendText("Hello World!");
            document.SaveToFile("Sample.doc", FileFormat.Doc);
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            Document document = new Document();
            document.LoadFromFile(@"C:\inetpub\wwwroot\temp1.docx");
            Section section = document.Sections[0];
            Paragraph para1 = section.Paragraphs[0];
            para1.Text = "Spire.Doc for .NET Introduction";
            document.SaveToFile(@"C:\inetpub\wwwroot\temp1.docx", FileFormat.Docx);
            return View();
        }
    }
}