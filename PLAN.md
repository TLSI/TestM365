# Plan: TestM365 — Plataforma de pruebas de integración M365

## Contexto

Se necesita una plataforma de demostración y pruebas para todas las integraciones Microsoft 365 previstas en **forlopd-pro**, con el fin de:
- Validar el comportamiento real de cada punto de integración antes de integrarlo en producción
- Presentar al cliente la experiencia de usuario final
- Detectar limitaciones técnicas (cuotas, permisos, latencias) en un entorno controlado

La plataforma replica exactamente el mismo stack tecnológico de forlopd-pro: **Nx monorepo + Angular 21 + NestJS 11 + Prisma 6 + PostgreSQL + Redis + MinIO**.

---

## Arquitectura

```
TestM365 (Nx monorepo)
├── apps/api/          NestJS 11 — puerto 3000
├── apps/web/          Angular 21 — portal de administración/testing — puerto 4200
├── apps/portal/       Angular 21 — portal simulado del cliente — puerto 4300
├── libs/shared/types/ Interfaces TypeScript compartidas
├── libs/shared/dtos/  DTOs con class-validator compartidos
└── libs/ui/components/ Componentes Angular reutilizables
```

**Dos portales:**
- **`web` (Admin/Testing):** Panel interno donde se ejercitan y validan todos los puntos de integración M365.
- **`portal` (Cliente simulado):** Experiencia que vería el cliente final: documentos de SharePoint (vía proxy), calendario, dashboard.

---

## Stack (idéntico a forlopd-pro)

| Capa | Tecnología |
|------|-----------|
| Frontend | Angular 21, PrimeNG 21, Sakai-NG (tema Aura), Tailwind CSS 4, NGX-Translate |
| Backend | NestJS 11, Prisma 6, Passport JWT |
| BD | PostgreSQL 16 |
| Cache/Colas | Redis 7, BullMQ |
| Storage | MinIO |
| M365 | @microsoft/microsoft-graph-client 3+, @azure/msal-node 2+ |
| DevOps | Nx 22, Docker Compose, Jest, ESLint, Prettier |

---

## Módulos M365 a implementar

| # | Área | Funcionalidades a probar |
|---|------|--------------------------|
| 1 | **Auth Azure AD** | OAuth 2.0 flow, almacenamiento de tokens (cifrado), renovación automática, estados (ACTIVE/EXPIRED/REVOKED/ERROR) |
| 2 | **Calendario** | Sync inicial, ver eventos Outlook, crear evento → Outlook, editar/eliminar bidireccional, webhooks, delta sync, link de Teams |
| 3 | **Correo** | Enviar email vía Outlook (Graph /me/sendMail), vincular email a cliente, ver bandeja, detección automática |
| 4 | **Documentos/SharePoint** | FORLOPD como proxy (client_credentials), crear site por cliente, carpetas por servicio, subir/listar/descargar, preview embebida (iframe), edición Office Online (link temporal anónimo) |
| 5 | **Teams** | Notificaciones a canal Teams (webhook entrante), crear reunión Teams |

---

## Decisiones de arquitectura M365

1. **Auth delegada (usuarios internos):** OAuth 2.0, tokens cifrados AES-256 en `MicrosoftConnection`.
2. **Auth de aplicación (SharePoint proxy):** `client_credentials` cacheado en Redis (TTL 58 min). Sin guest users.
3. **Webhooks de Graph:** `/webhooks/microsoft` valida `clientState` y procesa change notifications.
4. **Preview documentos:** Graph preview API → `<iframe>`. Sin cuenta M365 del cliente.
5. **Edición Office Online:** `createLink` con `type: edit, scope: anonymous`. Sin guest users.

---

## Variables de entorno (.env)

```env
DATABASE_URL=postgresql://testm365:testm365_dev@localhost:5432/testm365_dev
REDIS_HOST=localhost
REDIS_PORT=6379
MINIO_ENDPOINT=localhost
MINIO_PORT=9000
MINIO_ACCESS_KEY=minio_admin
MINIO_SECRET_KEY=minio_password
JWT_SECRET=...
JWT_REFRESH_SECRET=...

# Azure AD — añadir cuando estén disponibles
AZURE_CLIENT_ID=
AZURE_CLIENT_SECRET=
AZURE_TENANT_ID=
AZURE_REDIRECT_URI=http://localhost:3000/auth/microsoft/callback
SHAREPOINT_TENANT_DOMAIN=
```

El módulo M365 detecta si `AZURE_CLIENT_ID` está vacío y deshabilita las rutas OAuth mostrando un aviso en la UI.

---

## Fases

| Fase | Contenido | Días estimados |
|------|-----------|----------------|
| **1** | Setup monorepo Nx, apps, libs, Docker | 1-2 |
| **2** | Infra base: Prisma schema, auth JWT, Angular layouts | 2-4 |
| **3** | Azure AD OAuth + MicrosoftConnection | 4-6 |
| **4** | Calendario (Graph, sync, webhooks) | 6-9 |
| **5** | Correo (send, link, list) | 9-11 |
| **6** | Documentos/SharePoint (proxy, preview, edición) | 11-15 |
| **7** | Teams (notificaciones, reuniones) | 15-16 |

---

## Verificación final

- `npm run dev` → api:3000, web:4200, portal:4300
- Swagger en `http://localhost:3000/api/docs`
- Demo completo: conectar M365 → crear evento → subir doc → verlo en portal cliente → enviar email → notificación Teams
