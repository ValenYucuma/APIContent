// Los using se generan automáticamente, pero es bueno verificarlos
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

var builder = WebApplication.CreateBuilder(args);

// Agrega los servicios al contenedor.
// Habilita los controladores para que tu API los reconozca.
builder.Services.AddControllers();
// Habilita la exploración de endpoints para que Swagger los encuentre.
builder.Services.AddEndpointsApiExplorer();
// Genera la documentación de Swagger.
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configura el pipeline de solicitudes HTTP.

// En el entorno de desarrollo, habilita Swagger y su UI.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

// Habilita el enrutamiento. Esto es necesario para que el API detecte las rutas.
app.UseRouting();

app.UseAuthorization();

// ¡Esta es la línea clave que faltaba!
// Le dice al API que mapee las rutas de todos tus controladores.
app.MapControllers();

app.Run();
