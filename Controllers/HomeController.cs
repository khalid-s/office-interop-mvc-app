using DocumentGenerator.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

using Microsoft.Office.Interop.Word;

namespace DocumentGenerator.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            Application wordApplication = new Application();
            Document doc = wordApplication.Documents.Add();

            Paragraph para = doc.Paragraphs.Add();
            para.Range.Text = "Hello, this is a Word document generated using Interop.";

            // Save the document
            doc.SaveAs("GeneratedDocument.doc");

            // Close the application
            wordApplication.Quit();

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}