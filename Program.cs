using Microsoft.AspNetCore.Authentication.Cookies;

var builder = WebApplication.CreateBuilder(args);

// Добавление сервисов MVC
builder.Services.AddControllersWithViews();

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme).AddCookie(option=> {

    option.LoginPath = "/Access/Login";
    option.ExpireTimeSpan = TimeSpan.FromMinutes(20);
    option.AccessDeniedPath = new PathString("/excel/vipisky");
    

});

builder.Services.AddRazorPages(); // Добавление сервисов Razor страниц

var app = builder.Build();

// Использование статических файлов
app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();
// Защита от поддельных ресурсов
app.UseAntiforgery();

// Настройка маршрутизации для контроллеров MVC
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Access}/{action=Login}/{id?}");



app.Run();