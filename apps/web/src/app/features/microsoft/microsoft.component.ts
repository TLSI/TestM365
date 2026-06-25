import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TagModule } from 'primeng/tag';
import { DividerModule } from 'primeng/divider';
import { ToggleSwitchModule } from 'primeng/toggleswitch';
import { FormsModule } from '@angular/forms';
import { MessageModule } from 'primeng/message';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, TagModule, DividerModule, ToggleSwitchModule, MessageModule],
  template: `
    <div class="p-6 max-w-3xl">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Conexión Microsoft 365</h2>
        <p class="text-gray-500 mt-1">Gestiona la conexión OAuth con Azure AD</p>
      </div>

      @if (!status()?.configured) {
        <p-message
          severity="warn"
          text="Azure AD no está configurado. Añade AZURE_CLIENT_ID, AZURE_CLIENT_SECRET y AZURE_TENANT_ID al archivo .env y reinicia la API."
          styleClass="w-full mb-4"
        />
      }

      <p-card>
        <ng-template pTemplate="content">
          <div class="flex items-center justify-between mb-4">
            <div>
              <h3 class="font-semibold text-gray-900">Estado de la conexión</h3>
              @if (status()?.connected) {
                <p class="text-sm text-gray-500">Conectado como {{ status()?.email }}</p>
                <p class="text-xs text-gray-400">Token expira: {{ status()?.tokenExpiresAt | date:'medium' }}</p>
              } @else {
                <p class="text-sm text-gray-500">No conectado</p>
              }
            </div>
            <p-tag
              [value]="status()?.connected ? (status()?.status ?? 'ACTIVE') : 'Desconectado'"
              [severity]="status()?.connected ? (status()?.status === 'ACTIVE' ? 'success' : 'warn') : 'danger'"
            />
          </div>

          @if (status()?.connected) {
            <p-divider />
            <h4 class="font-medium text-gray-700 mb-3">Preferencias de sincronización</h4>
            <div class="space-y-3">
              <div class="flex items-center justify-between">
                <label class="text-sm">Sincronizar Calendario</label>
                <p-toggleswitch [(ngModel)]="syncCalendar" />
              </div>
              <div class="flex items-center justify-between">
                <label class="text-sm">Sincronizar Correo</label>
                <p-toggleswitch [(ngModel)]="syncMail" />
              </div>
              <div class="flex items-center justify-between">
                <label class="text-sm">Sincronizar Documentos</label>
                <p-toggleswitch [(ngModel)]="syncDocuments" />
              </div>
            </div>
            <p-divider />
            <p-button
              label="Desconectar Microsoft 365"
              icon="pi pi-times"
              severity="danger"
              [outlined]="true"
              [loading]="loading()"
              (onClick)="disconnect()"
            />
          } @else {
            <p-button
              label="Conectar con Microsoft 365"
              icon="pi pi-microsoft"
              [disabled]="!status()?.configured"
              [loading]="loading()"
              (onClick)="connect()"
            />
          }
        </ng-template>
      </p-card>

      <!-- Permisos que se solicitan -->
      <p-card styleClass="mt-4">
        <ng-template pTemplate="header">
          <div class="px-4 pt-4">
            <h3 class="font-semibold text-gray-900">Permisos OAuth solicitados</h3>
          </div>
        </ng-template>
        <ng-template pTemplate="content">
          <div class="grid grid-cols-2 gap-2">
            @for (scope of scopes; track scope.name) {
              <div class="flex items-start gap-2 text-sm">
                <i class="pi pi-check-circle text-green-500 mt-0.5 flex-shrink-0"></i>
                <div>
                  <span class="font-mono text-xs bg-gray-100 px-1 rounded">{{ scope.name }}</span>
                  <p class="text-gray-500 text-xs">{{ scope.desc }}</p>
                </div>
              </div>
            }
          </div>
        </ng-template>
      </p-card>
    </div>
  `,
})
export class MicrosoftComponent implements OnInit {
  private http = inject(HttpClient);

  status = signal<Record<string, unknown> | null>(null);
  loading = signal(false);
  syncCalendar = true;
  syncMail = true;
  syncDocuments = true;

  scopes = [
    { name: 'User.Read', desc: 'Perfil del usuario' },
    { name: 'Calendars.ReadWrite', desc: 'Calendario de Outlook' },
    { name: 'Mail.ReadWrite', desc: 'Leer correos' },
    { name: 'Mail.Send', desc: 'Enviar correos' },
    { name: 'Files.ReadWrite.All', desc: 'OneDrive/SharePoint' },
    { name: 'Sites.ReadWrite.All', desc: 'Sites de SharePoint' },
    { name: 'offline_access', desc: 'Refresh tokens' },
  ];

  ngOnInit() {
    this.loadStatus();
  }

  loadStatus() {
    this.http.get(`${environment.apiUrl}/auth/microsoft/status`).subscribe({
      next: (r: Record<string, unknown>) => {
        const data = (r as { data: Record<string, unknown> }).data;
        this.status.set(data);
        if (data) {
          this.syncCalendar = data['syncCalendar'] as boolean ?? true;
          this.syncMail = data['syncMail'] as boolean ?? true;
          this.syncDocuments = data['syncDocuments'] as boolean ?? true;
        }
      },
    });
  }

  connect() {
    this.loading.set(true);
    this.http.get<{ data: { url: string } }>(`${environment.apiUrl}/auth/microsoft/url`).subscribe({
      next: ({ data }) => {
        window.location.href = data.url;
      },
      error: () => this.loading.set(false),
    });
  }

  disconnect() {
    this.loading.set(true);
    this.http.delete(`${environment.apiUrl}/auth/microsoft/disconnect`).subscribe({
      next: () => { this.loadStatus(); this.loading.set(false); },
      error: () => this.loading.set(false),
    });
  }
}
