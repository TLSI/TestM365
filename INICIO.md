# TestM365 — Guía de arranque

Plataforma de pruebas de integración Microsoft 365, compuesta por un API NestJS y dos portales Angular.

---

## Requisitos previos

| Herramienta | Versión mínima | Verificar |
|-------------|---------------|-----------|
| Node.js | 20 LTS | `node --version` |
| npm | 10+ | `npm --version` |
| Docker Desktop | cualquiera | Debe estar en ejecución |
| Git | cualquiera | — |

---

## Primera puesta en marcha

Sigue los pasos en orden. Solo es necesario hacerlos una vez.

### Paso 1 — Instalar dependencias

```bash
npm install --legacy-peer-deps
```

### Paso 2 — Crear el archivo de entorno

Copia el archivo de ejemplo y revisa los valores (ya están configurados para desarrollo local):

```bash
cp .env.example .env
```

> El archivo `.env` ya existe si clonaste el repo con él incluido. En ese caso, omite este paso.

### Paso 3 — Arrancar Docker Desktop

Abre Docker Desktop y espera a que el motor esté activo (icono verde en la barra del sistema).

Luego levanta los contenedores:

```bash
docker compose up -d
```

Servicios que se levantan:

| Contenedor | Puerto local | Descripción |
|------------|-------------|-------------|
| `testm365-postgres` | 5433 | PostgreSQL 16 |
| `testm365-redis` | 6380 | Redis 7 |
| `testm365-minio` | 9002 / 9003 | MinIO (almacenamiento de objetos) |

Verifica que estén sanos:

```bash
docker compose ps
```

Todos deben aparecer en estado `healthy` o `running`.

### Paso 4 — Crear la base de datos (migración)

```bash
npm run prisma:migrate -- --name init
```

Esto crea todas las tablas en PostgreSQL.

### Paso 5 — Cargar datos de prueba (seed)

```bash
npm run prisma:seed
```

Crea los usuarios y el cliente de prueba:

| Email | Contraseña | Rol | Portal |
|-------|-----------|-----|--------|
| `admin@testm365.local` | `Admin1234!` | ADMIN | Web (admin) |
| `cliente@testm365.local` | `Admin1234!` | CLIENT | Portal (cliente) |

También crea un cliente de prueba llamado **ACME SL** vinculado al usuario cliente.

---

## Uso diario

Una vez completada la primera puesta en marcha, solo necesitas:

```bash
# 1. Asegúrate de que Docker está corriendo
docker compose up -d

# 2. Arrancar los tres servicios en paralelo
npm run dev
```

El comando `npm run dev` levanta simultáneamente:

| Servicio | URL | Descripción |
|---------|-----|-------------|
| API NestJS | http://localhost:3000 | Backend REST |
| Swagger UI | http://localhost:3000/api/docs | Documentación interactiva de la API |
| Portal web (admin) | http://localhost:4200 | Panel de testing de integración M365 |
| Portal cliente | http://localhost:4300 | Experiencia simulada del cliente |

---

## Parar el sistema

```bash
# Para los servidores Node: Ctrl+C en la terminal donde ejecutaste npm run dev

# Para los contenedores Docker:
docker compose stop

# Para eliminar los contenedores Y los datos (destructivo):
docker compose down -v
```

---

## Activar la integración con Microsoft 365

Por defecto el sistema arranca sin credenciales de Azure AD. Todas las rutas que no necesiten M365 funcionan con normalidad. Los módulos M365 muestran un aviso en la UI indicando que Azure AD no está configurado.

Cuando dispongas de credenciales, edita el archivo `.env`:

```env
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=tu_secreto_de_aplicacion
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_REDIRECT_URI=http://localhost:3000/auth/microsoft/callback

# Dominio del tenant de SharePoint (sin https://)
SHAREPOINT_TENANT_DOMAIN=tudespacho.sharepoint.com
```

Reinicia la API (`Ctrl+C` y `npm run dev`) y las funciones M365 se activarán automáticamente sin cambios de código.

### Permisos OAuth necesarios en el registro de la app en Azure

En **portal.azure.com → Registros de aplicaciones → tu app → Permisos de API**, añade permisos delegados de Microsoft Graph:

| Permiso | Para qué |
|---------|----------|
| `User.Read` | Perfil del usuario conectado |
| `Calendars.ReadWrite` | Leer y crear eventos de Outlook |
| `Mail.ReadWrite` | Leer bandeja de entrada |
| `Mail.Send` | Enviar correos vía Outlook |
| `Files.ReadWrite.All` | OneDrive / SharePoint |
| `Sites.ReadWrite.All` | Sites de SharePoint (proxy de documentos) |
| `offline_access` | Refresh tokens |

Y permisos de **aplicación** (para el proxy SharePoint sin usuario):

| Permiso | Para qué |
|---------|----------|
| `Sites.ReadWrite.All` | Proxy de documentos con `client_credentials` |
| `Files.ReadWrite.All` | Subir y listar ficheros en SharePoint |

En **Autenticación**, añade la URI de redirección:
```
http://localhost:3000/auth/microsoft/callback
```

---

## Herramientas adicionales

### Prisma Studio (explorador visual de la BD)

```bash
npm run prisma:studio
```

Abre un explorador en http://localhost:5555 donde puedes ver y editar registros directamente.

### Consola de Redis

```bash
docker exec -it testm365-redis redis-cli
```

### Consola de MinIO

Abre http://localhost:9003 en el navegador.
- Usuario: `minio_admin`
- Contraseña: `minio_password`

### Logs de los contenedores

```bash
docker compose logs -f postgres
docker compose logs -f redis
```

---

## Solución de problemas frecuentes

### Error: `connect ECONNREFUSED 127.0.0.1:5433`
Docker no está corriendo o los contenedores no están levantados. Ejecuta `docker compose up -d` y espera a que estén `healthy`.

### Error al migrar: `P1001 Can't reach database server`
Misma causa. También verifica que el puerto 5433 no esté ocupado por otro proceso:
```bash
# Windows
netstat -ano | findstr 5433
```

### La API arranca pero devuelve 401 en todas las rutas
El JWT_SECRET en `.env` debe tener al menos 32 caracteres. Verifica que el archivo `.env` existe en la raíz del proyecto.

### El portal web muestra "Azure AD no está configurado"
Es el comportamiento esperado cuando `AZURE_CLIENT_ID` está vacío en `.env`. El resto de funcionalidades (login, seed, CRUD) funcionan con normalidad.

### `npm install` falla con conflictos de peer dependencies
Usa siempre `--legacy-peer-deps`:
```bash
npm install --legacy-peer-deps
```

---

## Estructura del proyecto

```
TestM365/
├── apps/
│   ├── api/              NestJS 11 — API REST + módulos M365
│   │   └── prisma/       Schema y seed de la base de datos
│   ├── web/              Angular 21 — Portal admin/testing (:4200)
│   └── portal/           Angular 21 — Portal cliente simulado (:4300)
├── libs/
│   ├── shared/types/     Interfaces TypeScript compartidas
│   ├── shared/dtos/      DTOs con validación compartidos
│   └── ui/components/    Componentes Angular reutilizables
├── docker-compose.yml    PostgreSQL, Redis, MinIO
├── .env                  Variables de entorno (no en git)
├── .env.example          Plantilla de variables de entorno
└── INICIO.md             Este documento
```
