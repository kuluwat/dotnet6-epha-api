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
        [HttpPost("register_account", Name = "register_account")]
        public string register_account(RegisterAccountModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.register_account(param);

        }
        [HttpPost("update_register_account", Name = "update_register_account")]
        public string update_register_account(RegisterAccountModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.update_register_account(param);

        }
    }
}
