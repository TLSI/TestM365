import { Component, inject, signal, computed } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { ButtonModule } from 'primeng/button';
import { ToastModule } from 'primeng/toast';
import { DividerModule } from 'primeng/divider';
import { InputTextModule } from 'primeng/inputtext';
import { TextareaModule } from 'primeng/textarea';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';
import { firstValueFrom } from 'rxjs';

type Category = 'auth' | 'calendar' | 'mail' | 'documents' | 'teams';
type TestStatus = 'idle' | 'running' | 'success' | 'error';
type HttpMethod = 'GET' | 'POST' | 'DELETE';

interface ParamDef {
  key: string;
  label: string;
  defaultValue: string;
  placeholder?: string;
  multiline?: boolean;
}

interface TestDef {
  id: string;
  category: Category;
  name: string;
  description: string;
  m365Service: string;
  method: HttpMethod;
  endpoint: string;
  paramDefs: ParamDef[];
  bodyBuilder?: (params: Record<string, string>) => unknown;
  blobResponse?: boolean;
  dangerous?: boolean;
}

interface TestState {
  def: TestDef;
  status: TestStatus;
  result: unknown;
  error: string | null;
  durationMs: number | null;
  paramValues: Record<string, string>;
  showResult: boolean;
}

const h1 = () => new Date(Date.now() + 3600000).toISOString();
const h2 = () => new Date(Date.now() + 7200000).toISOString();

const TEST_DEFS: TestDef[] = [
  // ─── AUTH ──────────────────────────────────────────────────────────────────
  {
    id: 'auth-status',
    category: 'auth',
    name: 'Estado de conexión M365',
    description: 'Verifica si el usuario tiene tokens OAuth activos (MicrosoftConnection) en la BD.',
    m365Service: 'Azure AD',
    method: 'GET',
    endpoint: '/auth/microsoft/status',
    paramDefs: [],
  },
  {
    id: 'auth-url',
    category: 'auth',
    name: 'Obtener URL de autenticación OAuth',
    description: 'Genera la URL de login.microsoftonline.com con todos los scopes delegados para iniciar el flujo.',
    m365Service: 'Azure AD',
    method: 'GET',
    endpoint: '/auth/microsoft/url',
    paramDefs: [],
  },
  {
    id: 'auth-disconnect',
    category: 'auth',
    name: 'Desconectar cuenta M365',
    description: 'Elimina los tokens OAuth del usuario de la BD. Requiere reconexión para volver a usar servicios delegados.',
    m365Service: 'Azure AD',
    method: 'DELETE',
    endpoint: '/auth/microsoft/disconnect',
    paramDefs: [],
    dangerous: true,
  },

  // ─── CALENDAR ──────────────────────────────────────────────────────────────
  {
    id: 'cal-list',
    category: 'calendar',
    name: 'Listar eventos del calendario',
    description: 'Devuelve todos los eventos locales del usuario ordenados por fecha de inicio.',
    m365Service: 'Outlook Calendar',
    method: 'GET',
    endpoint: '/calendar/events',
    paramDefs: [],
  },
  {
    id: 'cal-sync',
    category: 'calendar',
    name: 'Sincronizar desde Outlook',
    description: 'Importa los próximos 90 días de eventos desde Graph /me/calendarView y hace upsert en la BD local.',
    m365Service: 'Outlook Calendar',
    method: 'POST',
    endpoint: '/calendar/sync',
    paramDefs: [],
    bodyBuilder: () => ({}),
  },
  {
    id: 'cal-create',
    category: 'calendar',
    name: 'Crear evento en Outlook',
    description: 'Crea un evento en la BD local y lo sincroniza con Outlook vía Graph /me/events. Con isOnline=true genera enlace Teams.',
    m365Service: 'Outlook Calendar + Teams',
    method: 'POST',
    endpoint: '/calendar/events',
    paramDefs: [
      { key: 'title', label: 'Título', defaultValue: 'Reunión de prueba TestM365', placeholder: 'Título del evento' },
      { key: 'description', label: 'Descripción', defaultValue: 'Evento creado desde el Test Suite M365', placeholder: 'Descripción' },
      { key: 'startTime', label: 'Inicio (ISO 8601)', defaultValue: h1(), placeholder: '2026-07-28T10:00:00.000Z' },
      { key: 'endTime', label: 'Fin (ISO 8601)', defaultValue: h2(), placeholder: '2026-07-28T11:00:00.000Z' },
      { key: 'location', label: 'Lugar (opcional)', defaultValue: 'Sala virtual', placeholder: 'Sala de reuniones' },
      { key: 'isOnline', label: '¿Online? Genera link Teams', defaultValue: 'false', placeholder: 'true / false' },
    ],
    bodyBuilder: (p) => ({
      title: p['title'],
      description: p['description'],
      startTime: p['startTime'],
      endTime: p['endTime'],
      location: p['location'],
      isOnline: p['isOnline'] === 'true',
      attendees: [],
    }),
  },
  {
    id: 'cal-delete',
    category: 'calendar',
    name: 'Eliminar evento (local + Outlook)',
    description: 'Elimina un evento de la BD local y de Outlook Calendar vía Graph DELETE /me/events/:externalId.',
    m365Service: 'Outlook Calendar',
    method: 'DELETE',
    endpoint: '/calendar/events/{eventId}',
    paramDefs: [
      { key: 'eventId', label: 'ID del evento local', defaultValue: '', placeholder: 'cuid del evento — obtenlo de "Listar eventos"' },
    ],
    dangerous: true,
  },

  // ─── MAIL ──────────────────────────────────────────────────────────────────
  {
    id: 'mail-inbox',
    category: 'mail',
    name: 'Bandeja de entrada de Outlook',
    description: 'Lista los mensajes más recientes del inbox vía Graph /me/mailFolders/inbox/messages.',
    m365Service: 'Outlook Mail',
    method: 'GET',
    endpoint: '/mail/inbox?top={top}',
    paramDefs: [
      { key: 'top', label: 'Nº de mensajes a traer', defaultValue: '5', placeholder: '5 (máx 50)' },
    ],
  },
  {
    id: 'mail-send',
    category: 'mail',
    name: 'Enviar email por Outlook',
    description: 'Envía un email usando la cuenta Outlook del usuario vía Graph /me/sendMail. Queda en la carpeta Enviados.',
    m365Service: 'Outlook Mail',
    method: 'POST',
    endpoint: '/mail/send',
    paramDefs: [
      { key: 'to', label: 'Destinatario', defaultValue: 'juan.perezlosa@gmail.com', placeholder: 'email@ejemplo.com' },
      { key: 'subject', label: 'Asunto', defaultValue: 'Test enviado desde TestM365', placeholder: 'Asunto del email' },
      { key: 'body', label: 'Cuerpo HTML', defaultValue: '<p>Email de prueba desde <strong>TestM365</strong>. Si recibes esto, la integración M365 funciona.</p>', placeholder: '<p>Contenido</p>', multiline: true },
    ],
    bodyBuilder: (p) => ({
      to: [p['to']],
      cc: [],
      subject: p['subject'],
      body: p['body'],
    }),
  },
  {
    id: 'mail-records',
    category: 'mail',
    name: 'Mis registros de email',
    description: 'Lista los últimos 50 registros EmailRecord del usuario almacenados localmente (enviados + vinculados).',
    m365Service: 'Outlook Mail',
    method: 'GET',
    endpoint: '/mail/records',
    paramDefs: [],
  },
  {
    id: 'mail-link',
    category: 'mail',
    name: 'Vincular email a cliente',
    description: 'Obtiene un mensaje de Outlook por su ID y crea un registro INBOUND vinculado al cliente en la BD.',
    m365Service: 'Outlook Mail',
    method: 'POST',
    endpoint: '/mail/link/{messageId}?clientId={clientId}',
    paramDefs: [
      { key: 'messageId', label: 'ID del mensaje Outlook', defaultValue: '', placeholder: 'id del mensaje — de "Bandeja de entrada"' },
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente — de "Listar clientes"' },
    ],
    bodyBuilder: () => ({}),
  },
  {
    id: 'mail-client',
    category: 'mail',
    name: 'Emails vinculados a un cliente',
    description: 'Lista todos los EmailRecord enlazados a un cliente específico (INBOUND + OUTBOUND).',
    m365Service: 'Outlook Mail',
    method: 'GET',
    endpoint: '/mail/records/client/{clientId}',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
    ],
  },

  // ─── DOCUMENTS / SHAREPOINT ────────────────────────────────────────────────
  {
    id: 'doc-clients',
    category: 'documents',
    name: '[ Helper ] Listar clientes',
    description: 'Lista todos los clientes de la BD. Útil para copiar un clientId válido antes de ejecutar los tests de SharePoint.',
    m365Service: 'FORLOPD (BD local)',
    method: 'GET',
    endpoint: '/clients',
    paramDefs: [],
  },
  {
    id: 'doc-site',
    category: 'documents',
    name: 'Crear site SharePoint para cliente',
    description: 'Crea un M365 Group + SharePoint Site con carpetas LOPD/Contratos/Correspondencia/Informes. Puede tardar hasta 60 segundos.',
    m365Service: 'SharePoint',
    method: 'POST',
    endpoint: '/documents/clients/{clientId}/site',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
    ],
    bodyBuilder: () => ({}),
  },
  {
    id: 'doc-list',
    category: 'documents',
    name: 'Listar documentos del cliente',
    description: 'Lista archivos del SharePoint Site del cliente vía Graph como proxy. El cliente nunca ve URLs de SharePoint.',
    m365Service: 'SharePoint',
    method: 'GET',
    endpoint: '/documents/clients/{clientId}?folder={folder}',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
      { key: 'folder', label: 'Subcarpeta (opcional)', defaultValue: '', placeholder: 'LOPD / Contratos / vacío = raíz' },
    ],
  },
  {
    id: 'doc-preview',
    category: 'documents',
    name: 'URL de vista previa (iframe)',
    description: 'Obtiene una URL de previsualización embebible para un documento. No requiere cuenta M365 del cliente.',
    m365Service: 'SharePoint / OneDrive',
    method: 'GET',
    endpoint: '/documents/clients/{clientId}/items/{itemId}/preview',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
      { key: 'itemId', label: 'ID del item SharePoint', defaultValue: '', placeholder: 'campo "id" de "Listar documentos"' },
    ],
  },
  {
    id: 'doc-editlink',
    category: 'documents',
    name: 'Enlace de edición Office Online',
    description: 'Genera un enlace temporal (4h) para editar en Word/Excel Online sin guest users en el tenant.',
    m365Service: 'SharePoint / Office Online',
    method: 'GET',
    endpoint: '/documents/clients/{clientId}/items/{itemId}/edit-link',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
      { key: 'itemId', label: 'ID del item SharePoint', defaultValue: '', placeholder: 'campo "id" de "Listar documentos"' },
    ],
  },
  {
    id: 'doc-download',
    category: 'documents',
    name: 'Descargar archivo (proxy)',
    description: 'Descarga el contenido del archivo desde SharePoint vía la API como proxy (el cliente nunca toca SharePoint directamente).',
    m365Service: 'SharePoint',
    method: 'GET',
    endpoint: '/documents/clients/{clientId}/items/{itemId}/download',
    paramDefs: [
      { key: 'clientId', label: 'ID del cliente', defaultValue: '', placeholder: 'cuid del cliente' },
      { key: 'itemId', label: 'ID del item SharePoint', defaultValue: '', placeholder: 'campo "id" de "Listar documentos"' },
    ],
    blobResponse: true,
  },

  // ─── TEAMS ─────────────────────────────────────────────────────────────────
  {
    id: 'teams-notify',
    category: 'teams',
    name: 'Enviar notificación a canal Teams',
    description: 'Envía un mensaje HTML a un canal vía Graph API (ChannelMessage.Send). La URL del canal incluye teamId y channelId.',
    m365Service: 'Microsoft Teams',
    method: 'POST',
    endpoint: '/teams/notify',
    paramDefs: [
      { key: 'channelUrl', label: 'URL del canal', defaultValue: '', placeholder: 'https://teams.microsoft.com/l/channel/...' },
      { key: 'title', label: 'Título', defaultValue: '🔔 Test desde TestM365', placeholder: 'Título del mensaje' },
      { key: 'subtitle', label: 'Subtítulo (opcional)', defaultValue: 'Sistema de validación FORLOPD', placeholder: 'Subtítulo' },
      { key: 'body', label: 'Mensaje', defaultValue: 'Esta es una notificación de prueba desde TestM365. Si aparece en el canal, la integración Teams funciona.', multiline: true },
    ],
    bodyBuilder: (p) => ({
      channelUrl: p['channelUrl'],
      title: p['title'],
      subtitle: p['subtitle'],
      body: p['body'],
      facts: [],
    }),
  },
  {
    id: 'teams-meeting',
    category: 'teams',
    name: 'Crear reunión Teams',
    description: 'Crea una reunión online vía Graph /me/onlineMeetings y devuelve el joinWebUrl listo para compartir.',
    m365Service: 'Microsoft Teams',
    method: 'POST',
    endpoint: '/teams/meeting',
    paramDefs: [
      { key: 'subject', label: 'Asunto', defaultValue: 'Reunión de prueba TestM365', placeholder: 'Asunto de la reunión' },
      { key: 'startTime', label: 'Inicio (ISO 8601)', defaultValue: h1(), placeholder: '2026-07-28T10:00:00.000Z' },
      { key: 'endTime', label: 'Fin (ISO 8601)', defaultValue: h2(), placeholder: '2026-07-28T11:00:00.000Z' },
    ],
    bodyBuilder: (p) => ({
      subject: p['subject'],
      startTime: p['startTime'],
      endTime: p['endTime'],
    }),
  },
];

@Component({
  selector: 'app-test-suite',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    ButtonModule,
    ToastModule,
    DividerModule,
    InputTextModule,
    TextareaModule,
  ],
  providers: [MessageService],
  template: `
    <p-toast />
    <div class="p-6 max-w-5xl">

      <!-- HEADER -->
      <div class="flex items-start justify-between mb-5">
        <div>
          <h2 class="text-2xl font-bold text-gray-900">Test Suite M365</h2>
          <p class="text-gray-500 mt-1 text-sm">
            Validación funcional de las {{ totalTests }} integraciones Microsoft 365 de FORLOPD
          </p>
        </div>
        <div class="flex items-center gap-2 shrink-0">
          <p-button
            label="Resetear"
            icon="pi pi-refresh"
            severity="secondary"
            size="small"
            [outlined]="true"
            [disabled]="runningAll()"
            (onClick)="resetAll()"
          />
          <p-button
            [label]="runningAll() ? 'Ejecutando...' : 'Ejecutar todos'"
            [icon]="runningAll() ? 'pi pi-spin pi-spinner' : 'pi pi-play'"
            size="small"
            [disabled]="runningAll()"
            (onClick)="runAll()"
          />
        </div>
      </div>

      <!-- STATS ROW -->
      <div class="grid grid-cols-4 gap-3 mb-6">
        <div class="bg-gray-50 border border-gray-200 rounded-lg p-3 text-center">
          <div class="text-2xl font-bold text-gray-800">{{ totalTests }}</div>
          <div class="text-xs text-gray-500 mt-0.5">Total</div>
        </div>
        <div class="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
          <div class="text-2xl font-bold text-green-600">{{ successCount() }}</div>
          <div class="text-xs text-green-600 mt-0.5">Pasados ✓</div>
        </div>
        <div class="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
          <div class="text-2xl font-bold text-red-600">{{ errorCount() }}</div>
          <div class="text-xs text-red-600 mt-0.5">Fallados ✗</div>
        </div>
        <div class="bg-gray-50 border border-gray-200 rounded-lg p-3 text-center">
          <div class="text-2xl font-bold text-gray-400">{{ pendingCount() }}</div>
          <div class="text-xs text-gray-400 mt-0.5">Pendientes</div>
        </div>
      </div>

      <!-- CATEGORY SECTIONS -->
      @for (cat of categories; track cat) {
        <div class="mb-8">

          <!-- Section header -->
          <div class="flex items-center gap-2 mb-3">
            <div [class]="'w-2.5 h-2.5 rounded-full shrink-0 ' + catDotClass(cat)"></div>
            <h3 class="font-semibold text-gray-600 text-xs uppercase tracking-wider">
              {{ catLabel(cat) }}
            </h3>
            <span class="text-xs text-gray-400">{{ testsForCat(cat).length }} tests</span>
            <div class="flex-1 h-px bg-gray-100 ml-1"></div>
          </div>

          <!-- Tests -->
          <div class="space-y-2">
            @for (st of testsForCat(cat); track st.def.id) {
              <div [class]="'rounded-lg border border-gray-200 overflow-hidden border-l-4 ' + catBorderClass(cat)">

                <!-- MAIN ROW -->
                <div class="flex items-center gap-2 px-3 py-2.5 bg-white">

                  <!-- Status icon -->
                  <div class="w-5 h-5 flex items-center justify-center shrink-0">
                    @switch (st.status) {
                      @case ('idle') {
                        <span class="w-2.5 h-2.5 rounded-full border-2 border-gray-300 block"></span>
                      }
                      @case ('running') {
                        <i class="pi pi-spin pi-spinner text-blue-500 text-sm"></i>
                      }
                      @case ('success') {
                        <i class="pi pi-check-circle text-green-500 text-sm"></i>
                      }
                      @case ('error') {
                        <i class="pi pi-times-circle text-red-500 text-sm"></i>
                      }
                    }
                  </div>

                  <!-- HTTP method badge -->
                  <span [class]="'text-xs font-mono font-bold px-1.5 py-0.5 rounded shrink-0 ' + methodClass(st.def.method)">
                    {{ st.def.method }}
                  </span>

                  <!-- Name + endpoint -->
                  <div class="flex-1 min-w-0">
                    <div class="flex items-center gap-2 flex-wrap">
                      <span class="font-semibold text-sm text-gray-900">{{ st.def.name }}</span>
                      @if (st.def.dangerous) {
                        <span class="text-xs text-orange-500 font-medium shrink-0">⚠ Destructiva</span>
                      }
                    </div>
                    <div class="text-xs text-gray-400 font-mono truncate">{{ st.def.endpoint }}</div>
                  </div>

                  <!-- M365 service badge (hidden on small screens) -->
                  <span class="text-xs text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full shrink-0 hidden lg:inline-block">
                    {{ st.def.m365Service }}
                  </span>

                  <!-- Duration -->
                  @if (st.durationMs !== null) {
                    <span class="text-xs text-gray-400 font-mono shrink-0">{{ st.durationMs }}ms</span>
                  }

                  <!-- Run button -->
                  <p-button
                    [label]="st.status === 'running' ? '' : 'Ejecutar'"
                    [icon]="st.status === 'running' ? 'pi pi-spin pi-spinner' : 'pi pi-play'"
                    size="small"
                    [outlined]="true"
                    [disabled]="st.status === 'running' || runningAll()"
                    [severity]="st.def.dangerous ? 'danger' : 'primary'"
                    (onClick)="runTest(st.def.id)"
                  />

                  <!-- Toggle result -->
                  @if (st.status === 'success' || st.status === 'error') {
                    <p-button
                      [icon]="st.showResult ? 'pi pi-chevron-up' : 'pi pi-chevron-down'"
                      size="small"
                      severity="secondary"
                      [text]="true"
                      (onClick)="toggleResult(st.def.id)"
                    />
                  }
                </div>

                <!-- PARAMS SECTION -->
                @if (st.def.paramDefs.length > 0) {
                  <div class="border-t border-gray-100 px-3 py-2.5 bg-gray-50">
                    <div class="text-xs font-semibold text-gray-400 uppercase tracking-wide mb-2">Parámetros</div>
                    <div [class]="st.def.paramDefs.length > 2 ? 'grid grid-cols-1 sm:grid-cols-2 gap-2' : 'grid grid-cols-1 gap-2'">
                      @for (p of st.def.paramDefs; track p.key) {
                        <div>
                          <label class="text-xs text-gray-600 block mb-1">{{ p.label }}</label>
                          @if (p.multiline) {
                            <textarea
                              pTextarea
                              [rows]="3"
                              [placeholder]="p.placeholder ?? ''"
                              [ngModel]="getParam(st.def.id, p.key)"
                              (ngModelChange)="setParam(st.def.id, p.key, $event)"
                              class="w-full text-xs font-mono"
                            ></textarea>
                          } @else {
                            <input
                              pInputText
                              size="small"
                              [placeholder]="p.placeholder ?? ''"
                              [ngModel]="getParam(st.def.id, p.key)"
                              (ngModelChange)="setParam(st.def.id, p.key, $event)"
                              class="w-full text-xs font-mono"
                            />
                          }
                        </div>
                      }
                    </div>
                  </div>
                }

                <!-- RESULT SECTION -->
                @if (st.showResult && (st.status === 'success' || st.status === 'error')) {
                  <div [class]="'border-t px-3 py-2.5 ' + (st.status === 'success' ? 'bg-green-50 border-green-200' : 'bg-red-50 border-red-200')">
                    <div class="flex items-center justify-between mb-2">
                      <span [class]="'text-xs font-semibold ' + (st.status === 'success' ? 'text-green-700' : 'text-red-700')">
                        {{ st.status === 'success' ? '✓ Respuesta exitosa' : '✗ Error' }}
                        @if (st.def.description) {
                          <span class="font-normal text-gray-500 ml-2">{{ st.def.description }}</span>
                        }
                      </span>
                      <p-button
                        label="Copiar"
                        icon="pi pi-copy"
                        size="small"
                        severity="secondary"
                        [text]="true"
                        (onClick)="copyResult(st)"
                      />
                    </div>
                    <pre [class]="'text-xs overflow-x-auto max-h-60 p-2 rounded font-mono leading-relaxed whitespace-pre-wrap break-all ' + (st.status === 'success' ? 'bg-green-100 text-green-900' : 'bg-red-100 text-red-900')">{{ formatResult(st) }}</pre>
                  </div>
                }

              </div>
            }
          </div>
        </div>
      }

    </div>
  `,
})
export class TestSuiteComponent {
  private http = inject(HttpClient);
  private msg = inject(MessageService);

  readonly categories: Category[] = ['auth', 'calendar', 'mail', 'documents', 'teams'];
  readonly totalTests = TEST_DEFS.length;

  states = signal<TestState[]>(
    TEST_DEFS.map((def) => ({
      def,
      status: 'idle' as TestStatus,
      result: null,
      error: null,
      durationMs: null,
      showResult: false,
      paramValues: Object.fromEntries(def.paramDefs.map((p) => [p.key, p.defaultValue])),
    }))
  );

  runningAll = signal(false);

  successCount = computed(() => this.states().filter((s) => s.status === 'success').length);
  errorCount = computed(() => this.states().filter((s) => s.status === 'error').length);
  pendingCount = computed(() => this.states().filter((s) => s.status === 'idle').length);

  testsForCat(cat: Category): TestState[] {
    return this.states().filter((s) => s.def.category === cat);
  }

  catLabel(cat: Category): string {
    return {
      auth: 'Autenticación Azure AD',
      calendar: 'Calendario Outlook',
      mail: 'Correo Outlook',
      documents: 'Documentos SharePoint',
      teams: 'Microsoft Teams',
    }[cat];
  }

  catDotClass(cat: Category): string {
    return {
      auth: 'bg-purple-400',
      calendar: 'bg-blue-400',
      mail: 'bg-amber-400',
      documents: 'bg-green-400',
      teams: 'bg-teal-400',
    }[cat];
  }

  catBorderClass(cat: Category): string {
    return {
      auth: 'border-l-purple-400',
      calendar: 'border-l-blue-400',
      mail: 'border-l-amber-400',
      documents: 'border-l-green-400',
      teams: 'border-l-teal-400',
    }[cat];
  }

  methodClass(method: HttpMethod): string {
    return {
      GET: 'bg-blue-100 text-blue-700',
      POST: 'bg-green-100 text-green-700',
      DELETE: 'bg-red-100 text-red-700',
    }[method];
  }

  getParam(testId: string, key: string): string {
    return this.states().find((s) => s.def.id === testId)?.paramValues[key] ?? '';
  }

  setParam(testId: string, key: string, value: string): void {
    this.states.update((states) =>
      states.map((s) =>
        s.def.id === testId ? { ...s, paramValues: { ...s.paramValues, [key]: value } } : s
      )
    );
  }

  toggleResult(testId: string): void {
    this.states.update((states) =>
      states.map((s) => (s.def.id === testId ? { ...s, showResult: !s.showResult } : s))
    );
  }

  private patch(testId: string, patch: Partial<TestState>): void {
    this.states.update((states) =>
      states.map((s) => (s.def.id === testId ? { ...s, ...patch } : s))
    );
  }

  private buildEndpoint(state: TestState): string {
    let ep = state.def.endpoint;
    ep = ep.replace(/\{(\w+)\}/g, (_, key) => encodeURIComponent(state.paramValues[key] ?? ''));
    // Remove query params with empty encoded values
    const qIdx = ep.indexOf('?');
    if (qIdx !== -1) {
      const path = ep.substring(0, qIdx);
      const kept = ep
        .substring(qIdx + 1)
        .split('&')
        .filter((kv) => {
          const eqIdx = kv.indexOf('=');
          return eqIdx === -1 || kv.substring(eqIdx + 1) !== '';
        })
        .join('&');
      return kept ? `${path}?${kept}` : path;
    }
    return ep;
  }

  async runTest(testId: string): Promise<void> {
    const state = this.states().find((s) => s.def.id === testId);
    if (!state) return;

    const start = Date.now();
    this.patch(testId, { status: 'running', result: null, error: null, showResult: false });

    const url = `${environment.apiUrl}${this.buildEndpoint(state)}`;
    const body = state.def.bodyBuilder?.(state.paramValues);

    try {
      let result: unknown;

      if (state.def.blobResponse) {
        const res = await firstValueFrom(
          this.http.get(url, { responseType: 'blob', observe: 'response' })
        );
        result = {
          size: res.body?.size,
          type: res.body?.type,
          contentDisposition: res.headers.get('content-disposition'),
          status: res.status,
        };
      } else if (state.def.method === 'GET') {
        result = await firstValueFrom(this.http.get(url));
      } else if (state.def.method === 'POST') {
        result = await firstValueFrom(this.http.post(url, body ?? {}));
      } else {
        result = await firstValueFrom(this.http.delete(url));
      }

      this.patch(testId, {
        status: 'success',
        result,
        durationMs: Date.now() - start,
        showResult: true,
      });
    } catch (e: unknown) {
      const err = e as { error?: { message?: string }; message?: string };
      this.patch(testId, {
        status: 'error',
        error: err.error?.message ?? err.message ?? 'Error desconocido',
        result: (e as { error?: unknown }).error ?? null,
        durationMs: Date.now() - start,
        showResult: true,
      });
    }
  }

  async runAll(): Promise<void> {
    this.runningAll.set(true);
    for (const state of this.states()) {
      if (state.def.dangerous) continue;
      await this.runTest(state.def.id);
      await new Promise<void>((r) => setTimeout(r, 300));
    }
    this.runningAll.set(false);
    this.msg.add({
      severity: this.errorCount() === 0 ? 'success' : 'warn',
      summary: `Test Suite completado: ${this.successCount()}/${this.totalTests} pasados`,
      life: 6000,
    });
  }

  resetAll(): void {
    this.states.update((states) =>
      states.map((s) => ({
        ...s,
        status: 'idle' as TestStatus,
        result: null,
        error: null,
        durationMs: null,
        showResult: false,
      }))
    );
  }

  formatResult(state: TestState): string {
    const data = state.result !== null ? state.result : state.error;
    try {
      return JSON.stringify(data, null, 2);
    } catch {
      return String(data);
    }
  }

  copyResult(state: TestState): void {
    void navigator.clipboard.writeText(this.formatResult(state));
    this.msg.add({ severity: 'info', summary: 'Copiado al portapapeles', life: 2000 });
  }
}
