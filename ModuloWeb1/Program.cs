using ModuloWeb.BROKER;

var builder = WebApplication.CreateBuilder(args);

// Probar conexión a MySQL
ConexionBD.ProbarConexion();

// Agregar servicios al contenedor
builder.Services.AddControllersWithViews();

var app = builder.Build();

// Configuración del pipeline HTTP
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

// Cambiar página inicial para que abra tu formulario de Orden de Compra
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=OrdenCompra}/{action=Crear}/{id?}");


app.Run();
