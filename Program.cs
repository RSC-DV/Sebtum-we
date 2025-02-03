using Microsoft.AspNetCore.Authentication.Cookies;

var builder = WebApplication.CreateBuilder(args);

// Добавление сервисов MVC
builder.Services.AddControllersWithViews();



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
    pattern: "{controller=Excel}/{action=Vipisky}/{id?}");



app.Run();