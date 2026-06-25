import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { DialogModule } from 'primeng/dialog';
import { InputTextModule } from 'primeng/inputtext';
import { TextareaModule } from 'primeng/textarea';
import { DatePickerModule } from 'primeng/datepicker';
import { ToggleSwitchModule } from 'primeng/toggleswitch';
import { TagModule } from 'primeng/tag';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

interface CalendarEvent {
  id: string;
  title: string;
  startTime: string;
  endTime: string;
  isOnline: boolean;
  meetingUrl?: string;
  source: string;
}

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, DialogModule,
    InputTextModule, TextareaModule, DatePickerModule, ToggleSwitchModule, TagModule, ToastModule],
  providers: [MessageService],
  template: `
    <div class="p-6">
      <p-toast />
      <div class="flex items-center justify-between mb-6">
        <div>
          <h2 class="text-2xl font-bold text-gray-900">Calendario</h2>
          <p class="text-gray-500 mt-1">Sincronización bidireccional con Outlook</p>
        </div>
        <div class="flex gap-2">
          <p-button label="Sincronizar desde Outlook" icon="pi pi-sync" severity="secondary" [loading]="syncing()" (onClick)="syncFromOutlook()" />
          <p-button label="Nuevo evento" icon="pi pi-plus" (onClick)="showDialog = true" />
        </div>
      </div>

      <p-card>
        <ng-template pTemplate="content">
          @if (loading()) {
            <div class="text-center py-8 text-gray-400"><i class="pi pi-spinner pi-spin text-2xl"></i></div>
          } @else if (events().length === 0) {
            <p class="text-center text-gray-400 py-8">Sin eventos. Pulsa "Sincronizar desde Outlook".</p>
          } @else {
            <div class="overflow-x-auto">
              <table class="w-full text-sm">
                <thead>
                  <tr class="border-b border-gray-200 text-left text-xs text-gray-500 uppercase tracking-wide">
                    <th class="pb-3 pr-4 font-medium">Título</th>
                    <th class="pb-3 pr-4 font-medium">Inicio</th>
                    <th class="pb-3 pr-4 font-medium">Fin</th>
                    <th class="pb-3 pr-4 font-medium">Online</th>
                    <th class="pb-3 pr-4 font-medium">Origen</th>
                    <th class="pb-3 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  @for (event of events(); track event.id) {
                    <tr class="border-b border-gray-100 hover:bg-gray-50">
                      <td class="py-3 pr-4 font-medium">{{ event.title }}</td>
                      <td class="py-3 pr-4 text-gray-600">{{ event.startTime | date:'short' }}</td>
                      <td class="py-3 pr-4 text-gray-600">{{ event.endTime | date:'short' }}</td>
                      <td class="py-3 pr-4">
                        @if (event.isOnline && event.meetingUrl) {
                          <a [href]="event.meetingUrl" target="_blank" class="text-blue-600 text-xs">Teams link</a>
                        } @else { <span class="text-gray-400">—</span> }
                      </td>
                      <td class="py-3 pr-4">
                        <p-tag [value]="event.source" [severity]="event.source === 'OUTLOOK' ? 'info' : 'success'" />
                      </td>
                      <td class="py-3">
                        <p-button icon="pi pi-trash" severity="danger" size="small" [text]="true" (onClick)="deleteEvent(event.id)" />
                      </td>
                    </tr>
                  }
                </tbody>
              </table>
            </div>
          }
        </ng-template>
      </p-card>

      <p-dialog header="Nuevo evento" [(visible)]="showDialog" [style]="{ width: '500px' }" [modal]="true">
        <div class="space-y-4">
          <div>
            <label class="text-sm font-medium">Título *</label>
            <input pInputText [(ngModel)]="newEvent.title" class="w-full mt-1" placeholder="Reunión de prueba" />
          </div>
          <div class="grid grid-cols-2 gap-4">
            <div>
              <label class="text-sm font-medium">Inicio *</label>
              <p-datepicker [(ngModel)]="newEvent.startDate" [showTime]="true" styleClass="w-full mt-1" />
            </div>
            <div>
              <label class="text-sm font-medium">Fin *</label>
              <p-datepicker [(ngModel)]="newEvent.endDate" [showTime]="true" styleClass="w-full mt-1" />
            </div>
          </div>
          <div>
            <label class="text-sm font-medium">Ubicación</label>
            <input pInputText [(ngModel)]="newEvent.location" class="w-full mt-1" />
          </div>
          <div class="flex items-center gap-3">
            <p-toggleswitch [(ngModel)]="newEvent.isOnline" />
            <label class="text-sm">Reunión online (genera link de Teams)</label>
          </div>
          <div>
            <label class="text-sm font-medium">Asistentes (emails, separados por coma)</label>
            <input pInputText [(ngModel)]="attendeesStr" class="w-full mt-1" placeholder="a@mail.com, b@mail.com" />
          </div>
        </div>
        <ng-template pTemplate="footer">
          <p-button label="Cancelar" severity="secondary" [text]="true" (onClick)="showDialog = false" />
          <p-button label="Crear y sincronizar con Outlook" icon="pi pi-send" [loading]="creating()" (onClick)="createEvent()" />
        </ng-template>
      </p-dialog>
    </div>
  `,
})
export class CalendarComponent implements OnInit {
  private http = inject(HttpClient);
  private msg = inject(MessageService);

  events = signal<CalendarEvent[]>([]);
  loading = signal(false);
  syncing = signal(false);
  creating = signal(false);
  showDialog = false;
  attendeesStr = '';
  newEvent = { title: '', startDate: new Date(), endDate: new Date(), location: '', isOnline: false };

  ngOnInit() { this.loadEvents(); }

  loadEvents() {
    this.loading.set(true);
    this.http.get<{ data: CalendarEvent[] }>(`${environment.apiUrl}/calendar/events`).subscribe({
      next: ({ data }) => { this.events.set(data); this.loading.set(false); },
      error: () => this.loading.set(false),
    });
  }

  syncFromOutlook() {
    this.syncing.set(true);
    this.http.post<{ data: { synced: number } }>(`${environment.apiUrl}/calendar/sync`, {}).subscribe({
      next: ({ data }) => {
        this.msg.add({ severity: 'success', summary: 'Sincronizado', detail: `${data.synced} eventos importados de Outlook` });
        this.syncing.set(false);
        this.loadEvents();
      },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.syncing.set(false); },
    });
  }

  createEvent() {
    this.creating.set(true);
    const payload = {
      title: this.newEvent.title,
      startTime: this.newEvent.startDate.toISOString(),
      endTime: this.newEvent.endDate.toISOString(),
      location: this.newEvent.location,
      isOnline: this.newEvent.isOnline,
      attendees: this.attendeesStr ? this.attendeesStr.split(',').map((s) => s.trim()) : [],
    };
    this.http.post(`${environment.apiUrl}/calendar/events`, payload).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Evento creado' }); this.showDialog = false; this.creating.set(false); this.loadEvents(); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.creating.set(false); },
    });
  }

  deleteEvent(id: string) {
    this.http.delete(`${environment.apiUrl}/calendar/events/${id}`).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Eliminado' }); this.loadEvents(); },
    });
  }
}
