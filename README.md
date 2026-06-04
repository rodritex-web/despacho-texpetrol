# Sistema de Despacho Texpetrol

Prototipo web responsivo para registrar cargas, confirmar checklist de seguridad, ingresar despachos y consultar historial visual.

## Abrir la app

Abrir `index.html` en Chrome, Edge o navegador Android.

## Tablas implementadas

### TablaDescargas

- Cliente
- Piloto
- Unidad
- Producto
- FechaDescarga
- HoraInicio
- HoraFin
- GalonesCargados
- GalonesRecibidos
- Diferencia
- Comentario
- Calificacion

### ChecklistPreSalida

- FechaCarga
- Piloto
- Unidad
- Planta
- Producto
- OdometroInicial
- OdometroFinal
- GalonesCargados
- EstadoViaje
- GalonesRestantes

## Pantalla Carga

Campos incluidos:

- Fecha Carga
- Planta: Amatitlan, Michatoya u Oakland
- Piloto
- Unidad asignada automaticamente por piloto, editable si se necesita cambiar
- Producto
- Odometro Inicial
- Galones Cargados
- Odometro Final

## Pantalla Ingrese Despacho

Al seleccionar un piloto, la app busca su carga activa en `ChecklistPreSalida`.

Si encuentra un viaje `EN RUTA`:

- Completa automaticamente la unidad.
- Completa automaticamente el producto.
- Muestra los galones pendientes.
- Coloca esos galones pendientes como `Galones Cargados` para el siguiente despacho.
- Permite agregar comentario del despacho.

Ejemplo: si el piloto cargo 3000 galones y entrega 2000, quedan 1000. En el siguiente despacho del mismo piloto, la pantalla muestra 1000 galones pendientes.

## Pantalla Calificar Despacho

Muestra los despachos registrados por cliente o identificador de despacho. Permite filtrar por cliente y por fecha de descarga.

Cada tarjeta incluye datos esenciales:

- Cliente
- ID despacho
- Fecha
- Piloto
- Unidad
- Producto
- Galones entregados
- Estado del despacho

La persona puede calificar el despacho con caritas: Malo, Regular, Bueno o Excelente. La calificacion se guarda en `TablaDescargas.Calificacion`.

Unidades por defecto:

- Cristian: TXT-19
- Mario: TXT-20
- Luis: TXT-21
- Carlos: TXT-22

## Lógica principal

Al guardar un despacho:

1. Calcula `Diferencia = GalonesCargados - GalonesRecibidos`.
2. Guarda el despacho en `TablaDescargas`.
3. Actualiza `ChecklistPreSalida.GalonesRestantes`.
4. Si `GalonesRestantes` es `0`, marca `EstadoViaje = FINALIZADO`.
5. Si queda pendiente, marca `EstadoViaje = EN RUTA`.

## Fechas

La fecha se captura con `input type="date"` y se guarda internamente como ISO 8601 usando medianoche UTC. Esto evita errores RFC3339 y permite migración futura a bases de datos, AppSheet, Power BI o APIs.

## Estados visuales

- Verde `#D4F4D2`: `Diferencia = 0`, muestra `FINALIZADO`.
- Amarillo `#FFF4C2`: `0 < Diferencia <= 500`, muestra `PENDIENTE GALONES: X`.
- Rojo `#F9D6D6`: `Diferencia > 500`, muestra `PENDIENTE GALONES: X`.

## Persistencia

Este prototipo guarda datos en `localStorage` del navegador bajo la clave `texpetrol-despacho-v1`.

Tambien incluye una pantalla de configuracion para preparar la migracion a SharePoint Lists. En esta version la configuracion queda guardada, pero la sincronizacion real aun requiere:

- URL del sitio de SharePoint
- Tenant ID
- Client ID de una app registrada en Microsoft Entra
- Permisos sobre las listas

Para producción, la idea es mover las tablas a una fuente centralizada como SharePoint, SQL, Dataverse o una API propia.
