namespace Sebtum.Models
{
    public class VMLogin
    {
        public string Email { get; set; }
        public string Password { get; set; }
        public bool KeepLoggedIn {  get; set; }
        public string Role { get; set; }
    }
}
