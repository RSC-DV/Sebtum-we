using Microsoft.AspNetCore.Mvc;

namespace Sebtum.Models
{
    public class ProductController : Controller
    {
        public IActionResult Index()
        {
            var products = new List<Product>
    {
        new Product { Id = 1, Name = "Телефон", Price = 500, Category = "Электроника" },
        new Product { Id = 2, Name = "Ноутбук \n Ноутбук \n </br>Ноутбук \n Ноутбук \n </br>Ноутбук \n Ноутбук \n </br>Ноутбук \n Ноутбук \n </br>", Price = 1000, Category = "Электроника" },
        new Product { Id = 3, Name = "Кроссовки", Price = 80, Category = "Одежда" },
        new Product { Id = 4, Name = "Шляпа", Price = 20, Category = "Аксессуары" },
        new Product { Id = 5, Name = "Кофеварка", Price = 150, Category = "Кухня" },
        new Product { Id = 1, Name = "Телефон", Price = 500, Category = "Электроника" },
        new Product { Id = 2, Name = "Ноутбук", Price = 1000, Category = "Электроника" },
        new Product { Id = 3, Name = "Кроссовки", Price = 80, Category = "Одежда" },
        new Product { Id = 4, Name = "Шляпа", Price = 20, Category = "Аксессуары" },
        new Product { Id = 5, Name = "Кофеварка", Price = 150, Category = "Кухня" },new Product { Id = 1, Name = "Телефон", Price = 500, Category = "Электроника" },
        new Product { Id = 2, Name = "Ноутбук", Price = 1000, Category = "Электроника" },
        new Product { Id = 3, Name = "Кроссовки", Price = 80, Category = "Одежда" },
        new Product { Id = 4, Name = "Шляпа", Price = 20, Category = "Аксессуары" },
        new Product { Id = 5, Name = "Кофеварка", Price = 150, Category = "Кухня" },
        new Product { Id = 1, Name = "Телефон", Price = 500, Category = "Электроника" },
        new Product { Id = 2, Name = "Ноутбук", Price = 1000, Category = "Электроника" },
        new Product { Id = 3, Name = "Кроссовки", Price = 80, Category = "Одежда" },
        new Product { Id = 4, Name = "Шляпа", Price = 20, Category = "Аксессуары" },
        new Product { Id = 5, Name = "Кофеварка", Price = 150, Category = "Кухня" },
        new Product { Id = 1, Name = "Телефон", Price = 500, Category = "Электроника" },
        new Product { Id = 2, Name = "Ноутбук", Price = 1000, Category = "Электроника" },
        new Product { Id = 3, Name = "Кроссовки", Price = 80, Category = "Одежда" },
        new Product { Id = 4, Name = "Шляпа", Price = 20, Category = "Аксессуары" },
        new Product { Id = 5, Name = "Кофеварка", Price = 150, Category = "Кухня" }
    };
            return View(products);
        }
    }
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Category { get; set; }
    }

}
