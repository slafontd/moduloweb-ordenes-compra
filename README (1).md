# Módulo de Órdenes de Compra – ModuloWeb

Este proyecto implementa un sistema web para generar órdenes de compra, almacenarlas en MySQL, generar un archivo Excel a partir de una plantilla y enviarlo por correo electrónico al proveedor mediante la API de SendGrid. Utiliza ASP.NET Core MVC, MySQL, ClosedXML y Docker para despliegue en Railway.

## Características principales

- Creación de órdenes de compra desde interfaz web.
- Selección de proveedor y productos filtrados por proveedor.
- Agregado y eliminación dinámica de productos en la orden.
- Cálculo automático del total.
- Generación de Excel desde plantilla corporativa.
- Envío automático del archivo Excel por correo.
- Configuración por variables de entorno segura.
- Trazabilidad de órdenes generadas.

## Arquitectura del Proyecto

La solución contiene cuatro capas:

### 1. **ModuloWeb1** (Web)
- Controladores MVC.
- Vistas Razor.
- Carpeta `Plantillas/` con la plantilla Excel `PlantillaOrdenes.xlsx`.

### 2. **ModuloWeb.MANAGER** (Lógica de negocio)
- Clase principal: `OrdenCompraManager`.
- Responsable de orquestar:
  - creación de órdenes,
  - generación del archivo Excel,
  - envío de correo.

### 3. **ModuloWeb.BROKER** (Acceso a datos)
- Clase principal: `OrdenCompraBroker`.
- Se encarga de toda la comunicación con MySQL.
- Obtiene proveedores, productos, inserta órdenes, detalles y trazabilidad.

### 4. **ModuloWeb.ENTITIES** (Modelos)
- Clases POCO para:
  - `Proveedor`
  - `Producto`
  - `OrdenCompra`
  - `DetalleOrden`

---

## Flujo de Funcionamiento

1. Usuario abre `/OrdenCompra/Crear`.
2. Selecciona proveedor y productos.
3. Opcional: puede escribir manualmente productos y precios.
4. JS agrega/elimina filas y recalcula total.
5. Al enviar:
   - Se guardan encabezado y detalles en BD.
   - Se genera Excel con ClosedXML.
   - Se envía correo al proveedor adjuntando la orden.
6. Se puede ver trazabilidad en `/OrdenCompra/Lista`.

---

## Base de Datos

Tablas utilizadas:

### **proveedores**
```
id (PK),
nombre,
nit,
correo,
telefono,
direccion
```

### **productos**
```
id (PK),
nombre,
precio,
id_proveedor (FK)
```

### **ordenes_compra**
```
id_orden (PK),
id_proveedor (FK),
total,
fecha,
estado
```

### **detalle_orden**
```
id_detalle (PK),
id_orden (FK),
id_producto (FK),
cantidad,
precio,
subtotal
```

---

## Generación del Excel

- Se carga la plantilla personalizada `PlantillaOrdenes.xlsx`.
- Se rellenan campos:
  - Datos del proveedor.
  - Fecha, moneda, condiciones de pago.
  - Líneas de productos (cantidad, descripción, precios, totales).
  - Subtotal y total final.
- Se guarda en:
```
/Ordenes/Orden_<id>.xlsx
```

---

## Envío del Correo (SendGrid)

Variables de entorno necesarias:

```
SENDGRID_API_KEY
FROM_EMAIL
ConnectionStrings__DefaultConnection
```

Código utiliza:

```csharp
var client = new SendGridClient(apiKey);
var msg = MailHelper.CreateSingleEmail(from, to, subject, body, null);
msg.AddAttachment("Orden.xlsx", base64);
```

---

## Variables de Entorno en Railway

| Variable | Descripción |
|---------|-------------|
| `ConnectionStrings__DefaultConnection` | Cadena de conexión MySQL |
| `FROM_EMAIL` | Email verificado en SendGrid |
| `SENDGRID_API_KEY` | API Key de SendGrid |
| `ASPNETCORE_ENVIRONMENT` | Development / Production |

---

## Docker y Railway

El proyecto se despliega con un **Dockerfile** que:

1. Usa .NET SDK para compilar.
2. Publica el proyecto web en `/app/publish`.
3. Usa imagen minimal de ASP.NET para ejecutar.
4. Railway expone el puerto automáticamente.

---

## Cómo Ejecutar Localmente

1. Configurar cadena de conexión en `appsettings.json`.
2. Restaurar paquetes:
```
dotnet restore
```
3. Ejecutar:
```
dotnet run --project ModuloWeb1
```

---

## Cómo Añadir Nuevos Campos

1. Agregar columna en BD.
2. Actualizar entidad en `ModuloWeb.ENTITIES`.
3. Ajustar consultas en **BROKER**.
4. Si aplica al Excel, modificar plantilla y código de `GenerarExcel()`.

---

## Autores
Santiago Lafont, Andrés Toro, Samuel Llano, Felipe Gómez.

Proyecto desarrollado como parte de un módulo de automatización de órdenes de compra con generación documental y envío automático.
