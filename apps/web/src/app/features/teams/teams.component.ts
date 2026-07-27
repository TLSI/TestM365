import { Component, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { TextareaModule } from 'primeng/textarea';
import { DatePickerModule } from 'primeng/datepicker';
import { DividerModule } from 'primeng/divider';
import { ToastModule } from 'primeng/toast';
import { MessageModule } from 'primeng/message';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, InputTextModule,
    TextareaModule, DatePickerModule, DividerModule, ToastModule, MessageModule],
  providers: [MessageService],
  template: `
    <div class="p-6 max-w-3xl">
      <p-toast />
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Microsoft Teams</h2>
        <p class="text-gray-500 mt-1">Notificaciones a canales y creación de reuniones</p>
      </div>

      <!-- Channel notification -->
      <p-card>
        <ng-template pTemplate="header">
          <div class="px-4 pt-4">
            <h3 class="font-semibold text-gray-900">Enviar notificación a canal Teams</h3>
            <p class="text-sm text-gray-500">Via Incoming Webhook del canal</p>
          </div>
        </ng-template>
        <ng-template pTemplate="content">
          <div class="space-y-4">
            <div>
              <label class="text-sm font-medium">URL del canal de Teams *</label>
              <input pInputText [(ngModel)]="notify.channelUrl" class="w-full mt-1" placeholder="https://teams.microsoft.com/l/channel/..." />
            </div>
            <div class="grid grid-cols-2 gap-4">
              <div>
                <label class="text-sm font-medium">Título *</label>
                <input pInputText [(ngModel)]="notify.title" class="w-full mt-1" placeholder="Alerta desde TestM365" />
              </div>
              <div>
                <label class="text-sm font-medium">Subtítulo</label>
                <input pInputText [(ngModel)]="notify.subtitle" class="w-full mt-1" placeholder="Sistema de pruebas" />
              </div>
            </div>
            <div>
              <label class="text-sm font-medium">Mensaje *</label>
              <textarea pTextarea [(ngModel)]="notify.body" rows="3" class="w-full mt-1" placeholder="Descripción de la notificación"></textarea>
            </div>
            <p-button label="Enviar notificación" icon="pi pi-send" [loading]="sendingNotify()" (onClick)="sendNotification()" />
          </div>
        </ng-template>
      </p-card>

      <p-divider />

      <!-- Create meeting -->
      <p-card>
        <ng-template pTemplate="header">
          <div class="px-4 pt-4">
            <h3 class="font-semibold text-gray-900">Crear reunión de Teams</h3>
            <p class="text-sm text-gray-500">Genera un enlace de reunión via Graph API</p>
          </div>
        </ng-template>
        <ng-template pTemplate="content">
          <div class="space-y-4">
            <div>
              <label class="text-sm font-medium">Asunto *</label>
              <input pInputText [(ngModel)]="meeting.subject" class="w-full mt-1" placeholder="Reunión de prueba Teams" />
            </div>
            <div class="grid grid-cols-2 gap-4">
              <div>
                <label class="text-sm font-medium">Inicio *</label>
                <p-datepicker [(ngModel)]="meeting.startDate" [showTime]="true" styleClass="w-full mt-1" />
              </div>
              <div>
                <label class="text-sm font-medium">Fin *</label>
                <p-datepicker [(ngModel)]="meeting.endDate" [showTime]="true" styleClass="w-full mt-1" />
              </div>
            </div>
            <p-button label="Crear reunión Teams" icon="pi pi-video" [loading]="creatingMeeting()" (onClick)="createMeeting()" />
          </div>

          @if (createdMeeting()) {
            <div class="mt-4 p-4 bg-green-50 rounded-lg">
              <p class="text-sm font-medium text-green-800">Reunión creada</p>
              <a [href]="createdMeeting()!['joinUrl']" target="_blank" class="text-blue-600 text-sm break-all">
                {{ createdMeeting()!['joinUrl'] }}
              </a>
            </div>
          }
        </ng-template>
      </p-card>
    </div>
  `,
})
export class TeamsComponent {
  private http = inject(HttpClient);
  private msg = inject(MessageService);

  sendingNotify = signal(false);
  creatingMeeting = signal(false);
  createdMeeting = signal<Record<string, unknown> | null>(null);

  notify = { channelUrl: '', title: 'Alerta desde TestM365', subtitle: 'Plataforma de pruebas M365', body: 'Esta es una notificación de prueba enviada desde TestM365.' };
  meeting = { subject: 'Reunión de prueba Teams', startDate: new Date(), endDate: new Date(Date.now() + 3600000) };

  sendNotification() {
    this.sendingNotify.set(true);
    this.http.post(`${environment.apiUrl}/teams/notify`, { ...this.notify }).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Notificación enviada al canal Teams' }); this.sendingNotify.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.sendingNotify.set(false); },
    });
  }

  createMeeting() {
    this.creatingMeeting.set(true);
    const payload = {
      subject: this.meeting.subject,
      startTime: this.meeting.startDate.toISOString(),
      endTime: this.meeting.endDate.toISOString(),
    };
    this.http.post<{ data: Record<string, unknown> }>(`${environment.apiUrl}/teams/meeting`, payload).subscribe({
      next: ({ data }) => { this.createdMeeting.set(data); this.msg.add({ severity: 'success', summary: 'Reunión de Teams creada' }); this.creatingMeeting.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.creatingMeeting.set(false); },
    });
  }
}
