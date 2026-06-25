import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { RouterLink } from '@angular/router';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TagModule } from 'primeng/tag';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, CardModule, ButtonModule, TagModule, RouterLink],
  template: `
    <div class="p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Dashboard — TestM365</h2>
        <p class="text-gray-500 mt-1">Estado general de la integración Microsoft 365</p>
      </div>

      <!-- Connection status -->
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
        <p-card>
          <ng-template pTemplate="content">
            <div class="flex items-center justify-between">
              <div>
                <p class="text-sm text-gray-500">Conexión M365</p>
                @if (msStatus()) {
                  <p-tag
                    [value]="msStatus()!.connected ? 'Conectado' : 'Desconectado'"
                    [severity]="msStatus()!.connected ? 'success' : 'danger'"
                  />
                  @if (msStatus()!.connected) {
                    <p class="text-xs text-gray-500 mt-1">{{ msStatus()!.email }}</p>
                  }
                } @else {
                  <p-tag value="Cargando..." severity="secondary" />
                }
              </div>
              <i class="pi pi-microsoft text-blue-600 text-3xl"></i>
            </div>
          </ng-template>
        </p-card>

        <p-card>
          <ng-template pTemplate="content">
            <div class="flex items-center justify-between">
              <div>
                <p class="text-sm text-gray-500">Azure AD configurado</p>
                <p-tag
                  [value]="msStatus()?.configured ? 'Sí' : 'No — falta .env'"
                  [severity]="msStatus()?.configured ? 'success' : 'warn'"
                />
              </div>
              <i class="pi pi-shield text-green-600 text-3xl"></i>
            </div>
          </ng-template>
        </p-card>

        <p-card>
          <ng-template pTemplate="content">
            <div class="flex items-center justify-between">
              <div>
                <p class="text-sm text-gray-500">Última sincronización</p>
                <p class="font-semibold">{{ msStatus()?.lastSyncAt ? (msStatus()!.lastSyncAt | date:'short') : 'Nunca' }}</p>
              </div>
              <i class="pi pi-sync text-purple-600 text-3xl"></i>
            </div>
          </ng-template>
        </p-card>
      </div>

      <!-- Quick actions -->
      <h3 class="text-lg font-semibold text-gray-900 mb-4">Paneles de prueba</h3>
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        @for (panel of panels; track panel.route) {
          <p-card styleClass="cursor-pointer hover:shadow-md transition-shadow">
            <ng-template pTemplate="content">
              <a [routerLink]="panel.route" class="flex items-start gap-4 no-underline">
                <div [class]="'w-12 h-12 rounded-xl flex items-center justify-center flex-shrink-0 ' + panel.bg">
                  <i [class]="'pi ' + panel.icon + ' text-white text-xl'"></i>
                </div>
                <div>
                  <p class="font-semibold text-gray-900">{{ panel.title }}</p>
                  <p class="text-sm text-gray-500 mt-1">{{ panel.desc }}</p>
                </div>
              </a>
            </ng-template>
          </p-card>
        }
      </div>
    </div>
  `,
})
export class DashboardComponent implements OnInit {
  private http = inject(HttpClient);
  msStatus = signal<Record<string, unknown> | null>(null);

  panels = [
    { title: 'Microsoft 365', icon: 'pi-microsoft', bg: 'bg-blue-600', route: '/microsoft', desc: 'Conectar / desconectar cuenta Azure AD. Ver estado y tokens.' },
    { title: 'Calendario', icon: 'pi-calendar', bg: 'bg-green-600', route: '/calendar', desc: 'Sincronizar eventos Outlook, crear reuniones con Teams.' },
    { title: 'Correo', icon: 'pi-envelope', bg: 'bg-orange-500', route: '/mail', desc: 'Enviar emails via Outlook, ver bandeja, vincular a cliente.' },
    { title: 'Documentos', icon: 'pi-file', bg: 'bg-purple-600', route: '/documents', desc: 'SharePoint proxy: subir, previsualizar, editar en Office Online.' },
    { title: 'Teams', icon: 'pi-comments', bg: 'bg-indigo-600', route: '/teams', desc: 'Notificaciones a canal Teams, crear reuniones Teams.' },
  ];

  ngOnInit() {
    this.http.get(`${environment.apiUrl}/auth/microsoft/status`).subscribe({
      next: (r: Record<string, unknown>) => this.msStatus.set((r as { data: Record<string, unknown> }).data),
      error: () => this.msStatus.set(null),
    });
  }
}
