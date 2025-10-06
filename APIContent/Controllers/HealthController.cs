using Microsoft.AspNetCore.Mvc;

namespace APIContent.Controllers
{
    [ApiController]
    [Route("/")]
    public class HealthController : ControllerBase
    {
        [HttpGet]
        public IActionResult Get()
        {
            return Ok(new
            {
                status = "running",
                environment = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "unknown",
                timeUtc = DateTime.UtcNow.ToString("o"),
                message = "APIContent is alive. Use /swagger to view API docs (if Swagger is enabled)."
            });
        }
    }
}
