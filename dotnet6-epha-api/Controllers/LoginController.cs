using Class;
using Microsoft.AspNetCore.Mvc;
using Model;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginController : ControllerBase
    {
        [HttpPost("check_authorization", Name = "check_authorization")]
        public string check_authorization(LoginUserModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.login(param);

        }
    }
}
