import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, RouterLink, CardModule, ButtonModule],
  template: `
    <div class="p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Bienvenido a su portal</h2>
        <p class="text-gray-500 mt-1">Acceda a sus documentos, citas y comunicaciones</p>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
        <p-card styleClass="cursor-pointer hover:shadow-md transition-shadow" routerLink="/documents">
          <ng-template pTemplate="content">
            <div class="flex items-center gap-4">
              <div class="w-12 h-12 bg-blue-100 rounded-xl flex items-center justify-center">
                <i class="pi pi-folder text-blue-600 text-xl"></i>
              </div>
              <div>
                <p class="font-semibold text-gray-900">Documentos</p>
                <p class="text-2xl font-bold text-blue-600">{{ docCount() }}</p>
                <p class="text-xs text-gray-400">archivos disponibles</p>
              </div>
            </div>
          </ng-template>
        </p-card>

        <p-card styleClass="cursor-pointer hover:shadow-md transition-shadow" routerLink="/calendar">
          <ng-template pTemplate="content">
            <div class="flex items-center gap-4">
              <div class="w-12 h-12 bg-green-100 rounded-xl flex items-center justify-center">
                <i class="pi pi-calendar text-green-600 text-xl"></i>
              </div>
              <div>
                <p class="font-semibold text-gray-900">Próximas citas</p>
                <p class="text-2xl font-bold text-green-600">{{ eventCount() }}</p>
                <p class="text-xs text-gray-400">próximos 30 días</p>
              </div>
            </div>
          </ng-template>
        </p-card>

        <p-card>
          <ng-template pTemplate="content">
            <div class="flex items-center gap-4">
              <div class="w-12 h-12 bg-purple-100 rounded-xl flex items-center justify-center">
                <i class="pi pi-microsoft text-purple-600 text-xl"></i>
              </div>
              <div>
                <p class="font-semibold text-gray-900">Plataforma</p>
                <p class="text-sm font-medium text-green-600 mt-1">Activa</p>
                <p class="text-xs text-gray-400">TestM365 Demo</p>
              </div>
            </div>
          </ng-template>
        </p-card>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <p-card>
          <ng-template pTemplate="header">
            <div class="px-4 pt-4 flex items-center justify-between">
              <h3 class="font-semibold text-gray-900">Acceso rápido</h3>
            </div>
          </ng-template>
          <ng-template pTemplate="content">
            <div class="space-y-2">
              <a routerLink="/documents" class="flex items-center gap-3 p-3 rounded-lg hover:bg-gray-50 cursor-pointer text-gray-700">
                <i class="pi pi-folder-open text-blue-500"></i>
                <span class="text-sm">Ver todos mis documentos</span>
              </a>
              <a routerLink="/calendar" class="flex items-center gap-3 p-3 rounded-lg hover:bg-gray-50 cursor-pointer text-gray-700">
                <i class="pi pi-calendar-plus text-green-500"></i>
                <span class="text-sm">Ver calendario de citas</span>
              </a>
            </div>
          </ng-template>
        </p-card>

        <p-card>
          <ng-template pTemplate="header">
            <div class="px-4 pt-4">
              <h3 class="font-semibold text-gray-900">Información</h3>
            </div>
          </ng-template>
          <ng-template pTemplate="content">
            <div class="space-y-2 text-sm text-gray-600">
              <p class="flex items-start gap-2">
                <i class="pi pi-shield text-blue-500 mt-0.5"></i>
                Sus documentos se almacenan de forma segura en SharePoint
              </p>
              <p class="flex items-start gap-2">
                <i class="pi pi-lock text-green-500 mt-0.5"></i>
                Acceso exclusivo — nunca comparte credenciales de M365
              </p>
              <p class="flex items-start gap-2">
                <i class="pi pi-sync text-purple-500 mt-0.5"></i>
                Sincronización automática con el equipo de gestión
              </p>
            </div>
          </ng-template>
        </p-card>
      </div>
    </div>
  `,
})
export class DashboardComponent implements OnInit {
  private http = inject(HttpClient);

  docCount = signal(0);
  eventCount = signal(0);

  ngOnInit() {
    this.http.get<{ data: Array<{ startTime: string }> }>(`${environment.apiUrl}/calendar/events`).subscribe({
      next: ({ data }) => {
        const now = new Date();
        const in30 = new Date(Date.now() + 30 * 24 * 3600000);
        this.eventCount.set(
          data.filter((e) => {
            const d = new Date(e.startTime);
            return d >= now && d <= in30;
          }).length
        );
      },
    });
  }
}
