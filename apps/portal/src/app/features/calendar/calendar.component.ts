import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TagModule } from 'primeng/tag';
import { environment } from '../../../environments/environment';

interface CalendarEvent {
  id: string;
  title: string;
  startTime: string;
  endTime: string;
  source: string;
  isOnline?: boolean;
  meetingUrl?: string;
}

@Component({
  standalone: true,
  imports: [CommonModule, CardModule, ButtonModule, TagModule],
  template: `
    <div class="p-6">
      <div class="flex items-center justify-between mb-6">
        <div>
          <h2 class="text-2xl font-bold text-gray-900">Mis Citas</h2>
          <p class="text-gray-500 mt-1">Próximas reuniones y eventos programados</p>
        </div>
        <p-button label="Actualizar" icon="pi pi-refresh" severity="secondary" [loading]="loading()" (onClick)="load()" />
      </div>

      @if (loading()) {
        <div class="text-center py-12 text-gray-400">
          <i class="pi pi-spinner pi-spin text-4xl mb-3 block"></i>
          <p>Cargando eventos...</p>
        </div>
      } @else if (upcoming().length === 0) {
        <div class="text-center py-12 text-gray-400">
          <i class="pi pi-calendar text-4xl mb-3 block text-gray-300"></i>
          <p class="font-medium">No hay citas próximas</p>
          <p class="text-sm mt-1">Contacte con su gestor para programar una reunión</p>
        </div>
      } @else {
        <div class="space-y-3">
          @for (event of upcoming(); track event.id) {
            <p-card>
              <ng-template pTemplate="content">
                <div class="flex items-start gap-4">
                  <!-- Date block -->
                  <div class="text-center bg-blue-50 rounded-xl p-3 min-w-16 flex-shrink-0">
                    <p class="text-2xl font-bold text-blue-600 leading-none">{{ event.startTime | date:'d' }}</p>
                    <p class="text-xs text-blue-500 uppercase font-medium">{{ event.startTime | date:'MMM' }}</p>
                  </div>

                  <!-- Details -->
                  <div class="flex-1">
                    <div class="flex items-center gap-2 mb-1">
                      <h3 class="font-semibold text-gray-900">{{ event.title }}</h3>
                      @if (event.isOnline) {
                        <p-tag value="Online" severity="info" />
                      }
                      <p-tag [value]="event.source" [severity]="event.source === 'OUTLOOK' ? 'success' : 'secondary'" />
                    </div>
                    <p class="text-sm text-gray-500">
                      <i class="pi pi-clock mr-1"></i>
                      {{ event.startTime | date:'HH:mm' }} – {{ event.endTime | date:'HH:mm' }}
                    </p>
                    @if (event.isOnline && event.meetingUrl) {
                      <a [href]="event.meetingUrl" target="_blank"
                        class="inline-flex items-center gap-1 mt-2 text-sm text-blue-600 hover:underline">
                        <i class="pi pi-video"></i> Unirse a la reunión de Teams
                      </a>
                    }
                  </div>
                </div>
              </ng-template>
            </p-card>
          }
        </div>
      }

      <!-- Past events section -->
      @if (past().length > 0) {
        <div class="mt-8">
          <h3 class="text-sm font-medium text-gray-400 mb-3 uppercase tracking-wide">Citas anteriores</h3>
          <div class="space-y-2">
            @for (event of past().slice(0, 5); track event.id) {
              <div class="flex items-center gap-3 p-3 bg-gray-50 rounded-lg opacity-60">
                <i class="pi pi-calendar-minus text-gray-400"></i>
                <span class="text-sm text-gray-600">{{ event.title }}</span>
                <span class="text-xs text-gray-400 ml-auto">{{ event.startTime | date:'dd/MM/yyyy' }}</span>
              </div>
            }
          </div>
        </div>
      }
    </div>
  `,
})
export class CalendarComponent implements OnInit {
  private http = inject(HttpClient);

  events = signal<CalendarEvent[]>([]);
  loading = signal(false);

  upcoming = signal<CalendarEvent[]>([]);
  past = signal<CalendarEvent[]>([]);

  ngOnInit() { this.load(); }

  load() {
    this.loading.set(true);
    this.http.get<{ data: CalendarEvent[] }>(`${environment.apiUrl}/calendar/events`).subscribe({
      next: ({ data }) => {
        const now = new Date();
        const sorted = [...data].sort((a, b) =>
          new Date(a.startTime).getTime() - new Date(b.startTime).getTime()
        );
        this.upcoming.set(sorted.filter((e) => new Date(e.startTime) >= now));
        this.past.set(sorted.filter((e) => new Date(e.startTime) < now).reverse());
        this.loading.set(false);
      },
      error: () => this.loading.set(false),
    });
  }
}
