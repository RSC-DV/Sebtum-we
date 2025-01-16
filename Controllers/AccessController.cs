using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Sebtum.Models;

namespace Sebtum.Controllers
{
    public class AccessController : Controller
    {
        // Метод отображения формы входа
        public IActionResult Login()
        {
            // Проверяем, аутентифицирован ли пользователь
            ClaimsPrincipal claimUser = HttpContext.User;
            if (claimUser.Identity.IsAuthenticated)
                // Если аутентифицирован, перенаправляем на главную страницу
                return RedirectToAction("Index", "Home");
            // В противном случае, отображаем форму входа
            return View();
        }

        // Метод обработки данных формы входа
        [HttpPost]
        public async Task<IActionResult> Login(VMLogin modelLogin)
        {
            if (modelLogin.Email == "user1@example.com" && modelLogin.Password == "12345")
            {
                // Создаем список атрибутов пользователя (Claims)
                List<Claim> claims = new List<Claim>() {
                    // Идентификатор пользователя (Email)
                    new Claim(ClaimTypes.NameIdentifier, modelLogin.Email),
                    // Дополнительный атрибут (Example Role)
                    new Claim(ClaimTypes.Role, "default"),
                    new Claim("Name", "John Doe"),
                    new Claim("id", "25", ClaimValueTypes.Integer)
                };

                // Создаем объект ClaimsIdentity, который хранит информацию о пользователе
                ClaimsIdentity claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

                // Создаем объект AuthenticationProperties, который хранит информацию о сессии пользователя
                AuthenticationProperties properties = new AuthenticationProperties()
                {
                    // Разрешить обновление токена (не используется в этом примере)
                    AllowRefresh = true,
                    // Запомнить пользователя, если он поставил галочку "Запомнить меня"
                    IsPersistent = modelLogin.KeepLoggedIn
                };

                // Аутентифицируем пользователя и записываем его в сессию
                await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), properties);

                // Перенаправляем на главную страницу
                return RedirectToAction("Index", "Home");
            }
            // Проверка введенных пользователем данных
            if (modelLogin.Email == "user@example.com" && modelLogin.Password == "12345")
            {
                // Создаем список атрибутов пользователя (Claims)
                List<Claim> claims = new List<Claim>() {
                    // Идентификатор пользователя (Email)
                    new Claim(ClaimTypes.NameIdentifier, modelLogin.Email),
                    // Дополнительный атрибут (Example Role)
                    new Claim(ClaimTypes.Role, "Admin"),
                    new Claim("Name", "Jane Doe"),
                    new Claim("id", "25", ClaimValueTypes.Integer)
                };

                // Создаем объект ClaimsIdentity, который хранит информацию о пользователе
                ClaimsIdentity claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

                // Создаем объект AuthenticationProperties, который хранит информацию о сессии пользователя
                AuthenticationProperties properties = new AuthenticationProperties()
                {
                    // Разрешить обновление токена (не используется в этом примере)
                    AllowRefresh = true,
                    // Запомнить пользователя, если он поставил галочку "Запомнить меня"
                    IsPersistent = modelLogin.KeepLoggedIn
                };

                // Аутентифицируем пользователя и записываем его в сессию
                await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), properties);

                // Перенаправляем на главную страницу
                return RedirectToAction("Index", "Home");
            }

            // Если введенные данные неверны, выводим сообщение об ошибке
            ViewData["ValidateMessage"] = "Пользователь не найден";
            return View();
        }

        // Метод обработки выхода пользователя из системы
        public async Task<IActionResult> LogOut()
        {
            // Удаляем данные пользователя из сессии
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            // Перенаправляем на страницу входа
            return RedirectToAction("Login", "Access");
        }
    }
}