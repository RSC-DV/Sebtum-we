using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Sebtum.Models;

using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using System.Security.Claims;
namespace Sebtum.Controllers;


public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {   string like = "s";
        ViewData["k"] = "Добро пожаловать на сайт!";
        ViewData["Message"] = "Это сообщение из контроллера";

        return View();
    }
     public ActionResult GetMessage()
    {
        return PartialView("_GetMessage");
    }
    public IActionResult SnakePage()
    {
        return View();
    }
     public string Index1(int a)
        {
            return "Привет"+a;
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
