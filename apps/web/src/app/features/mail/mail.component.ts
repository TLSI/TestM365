import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { DialogModule } from 'primeng/dialog';
import { InputTextModule } from 'primeng/inputtext';
import { TextareaModule } from 'primeng/textarea';
import { TabsModule } from 'primeng/tabs';
import { TagModule } from 'primeng/tag';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

interface InboxMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  from?: { emailAddress: { address: string } };
}

interface EmailRecord {
  id: string;
  direction: string;
  from?: string;
  to?: string[];
  subject: string;
  sentAt: string;
  clientId?: string;
}

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, DialogModule,
    InputTextModule, TextareaModule, TabsModule, TagModule, ToastModule],
  providers: [MessageService],
  template: `
    <div class="p-6">
      <p-toast />
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Correo Electrónico</h2>
        <p class="text-gray-500 mt-1">Envío y vinculación de emails via Outlook (Graph API)</p>
      </div>

      <p-tabs value="0">
        <p-tablist>
          <p-tab value="0">Bandeja de Outlook</p-tab>
          <p-tab value="1">Enviar email</p-tab>
          <p-tab value="2">Emails vinculados</p-tab>
        </p-tablist>

        <p-tabpanels>
          <p-tabpanel value="0">
            <div class="flex justify-end mb-3">
              <p-button label="Cargar bandeja de Outlook" icon="pi pi-refresh" severity="secondary" [loading]="loadingInbox()" (onClick)="loadInbox()" />
            </div>
            @if (inbox().length === 0) {
              <p class="text-center text-gray-400 py-8">Pulsa "Cargar bandeja" para ver los emails de Outlook</p>
            } @else {
              <div class="overflow-x-auto">
                <table class="w-full text-sm">
                  <thead>
                    <tr class="border-b border-gray-200 text-left text-xs text-gray-500 uppercase tracking-wide">
                      <th class="pb-3 pr-4 font-medium">De</th>
                      <th class="pb-3 pr-4 font-medium">Asunto</th>
                      <th class="pb-3 pr-4 font-medium">Vista previa</th>
                      <th class="pb-3 pr-4 font-medium">Fecha</th>
                      <th class="pb-3 font-medium">Vincular</th>
                    </tr>
                  </thead>
                  <tbody>
                    @for (msg of inbox(); track msg.id) {
                      <tr class="border-b border-gray-100 hover:bg-gray-50">
                        <td class="py-3 pr-4 text-xs text-gray-600">{{ msg.from ? msg.from.emailAddress.address : '' }}</td>
                        <td class="py-3 pr-4 font-medium max-w-xs truncate">{{ msg.subject }}</td>
                        <td class="py-3 pr-4 text-xs text-gray-400 max-w-xs truncate">{{ msg.bodyPreview }}</td>
                        <td class="py-3 pr-4 text-xs">{{ msg.receivedDateTime | date:'short' }}</td>
                        <td class="py-3">
                          <p-button label="Vincular" size="small" [text]="true" (onClick)="openLinkDialog(msg)" />
                        </td>
                      </tr>
                    }
                  </tbody>
                </table>
              </div>
            }
          </p-tabpanel>

          <p-tabpanel value="1">
            <div class="max-w-xl space-y-4 pt-2">
              <div>
                <label class="text-sm font-medium">Para (emails separados por coma) *</label>
                <input pInputText [(ngModel)]="sendForm.to" class="w-full mt-1" placeholder="destinatario@example.com" />
              </div>
              <div>
                <label class="text-sm font-medium">CC</label>
                <input pInputText [(ngModel)]="sendForm.cc" class="w-full mt-1" />
              </div>
              <div>
                <label class="text-sm font-medium">Asunto *</label>
                <input pInputText [(ngModel)]="sendForm.subject" class="w-full mt-1" />
              </div>
              <div>
                <label class="text-sm font-medium">Cuerpo (HTML) *</label>
                <textarea pTextarea [(ngModel)]="sendForm.body" rows="6" class="w-full mt-1"></textarea>
              </div>
              <p-button label="Enviar via Outlook" icon="pi pi-send" [loading]="sending()" (onClick)="sendEmail()" />
            </div>
          </p-tabpanel>

          <p-tabpanel value="2">
            @if (records().length === 0) {
              <p class="text-center text-gray-400 py-8">No hay emails vinculados aún</p>
            } @else {
              <div class="overflow-x-auto">
                <table class="w-full text-sm">
                  <thead>
                    <tr class="border-b border-gray-200 text-left text-xs text-gray-500 uppercase tracking-wide">
                      <th class="pb-3 pr-4 font-medium">Dirección</th>
                      <th class="pb-3 pr-4 font-medium">De/Para</th>
                      <th class="pb-3 pr-4 font-medium">Asunto</th>
                      <th class="pb-3 pr-4 font-medium">Fecha</th>
                      <th class="pb-3 font-medium">Cliente</th>
                    </tr>
                  </thead>
                  <tbody>
                    @for (r of records(); track r.id) {
                      <tr class="border-b border-gray-100 hover:bg-gray-50">
                        <td class="py-3 pr-4">
                          <p-tag [value]="r.direction" [severity]="r.direction === 'OUTBOUND' ? 'success' : 'info'" />
                        </td>
                        <td class="py-3 pr-4 text-xs text-gray-600">{{ r.direction === 'OUTBOUND' ? r.to?.join(', ') : r.from }}</td>
                        <td class="py-3 pr-4">{{ r.subject }}</td>
                        <td class="py-3 pr-4 text-xs">{{ r.sentAt | date:'short' }}</td>
                        <td class="py-3 text-xs text-gray-500">{{ r.clientId ?? '—' }}</td>
                      </tr>
                    }
                  </tbody>
                </table>
              </div>
            }
          </p-tabpanel>
        </p-tabpanels>
      </p-tabs>

      <p-dialog header="Vincular email a cliente" [(visible)]="showLinkDialog" [style]="{ width: '400px' }" [modal]="true">
        <div class="space-y-4">
          <p class="text-sm text-gray-600">Email: <strong>{{ selectedMsg()?.subject }}</strong></p>
          <div>
            <label class="text-sm font-medium">ID del cliente</label>
            <input pInputText [(ngModel)]="linkClientId" class="w-full mt-1" placeholder="ID del cliente en la BD" />
          </div>
        </div>
        <ng-template pTemplate="footer">
          <p-button label="Cancelar" severity="secondary" [text]="true" (onClick)="showLinkDialog = false" />
          <p-button label="Vincular" [loading]="linking()" (onClick)="linkEmail()" />
        </ng-template>
      </p-dialog>
    </div>
  `,
})
export class MailComponent implements OnInit {
  private http = inject(HttpClient);
  private msg = inject(MessageService);

  inbox = signal<InboxMessage[]>([]);
  records = signal<EmailRecord[]>([]);
  loadingInbox = signal(false);
  loadingRecords = signal(false);
  sending = signal(false);
  linking = signal(false);
  showLinkDialog = false;
  linkClientId = '';
  selectedMsg = signal<InboxMessage | null>(null);
  sendForm = { to: '', cc: '', subject: 'Prueba desde TestM365', body: '<p>Hola, este es un email de prueba enviado via Microsoft Graph API desde la plataforma TestM365.</p>' };

  ngOnInit() { this.loadRecords(); }

  loadInbox() {
    this.loadingInbox.set(true);
    this.http.get<{ data: InboxMessage[] }>(`${environment.apiUrl}/mail/inbox`).subscribe({
      next: ({ data }) => { this.inbox.set(data); this.loadingInbox.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.loadingInbox.set(false); },
    });
  }

  loadRecords() {
    this.loadingRecords.set(true);
    this.http.get<{ data: EmailRecord[] }>(`${environment.apiUrl}/mail/records`).subscribe({
      next: ({ data }) => { this.records.set(data); this.loadingRecords.set(false); },
      error: () => this.loadingRecords.set(false),
    });
  }

  sendEmail() {
    this.sending.set(true);
    const payload = {
      to: this.sendForm.to.split(',').map((s) => s.trim()),
      cc: this.sendForm.cc ? this.sendForm.cc.split(',').map((s) => s.trim()) : [],
      subject: this.sendForm.subject,
      body: this.sendForm.body,
    };
    this.http.post(`${environment.apiUrl}/mail/send`, payload).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Email enviado via Outlook' }); this.sending.set(false); this.loadRecords(); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.sending.set(false); },
    });
  }

  openLinkDialog(msg: InboxMessage) { this.selectedMsg.set(msg); this.showLinkDialog = true; }

  linkEmail() {
    const id = this.selectedMsg()?.id;
    if (!id) return;
    this.linking.set(true);
    this.http.post(`${environment.apiUrl}/mail/link/${id}?clientId=${this.linkClientId}`, {}).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Email vinculado' }); this.showLinkDialog = false; this.linking.set(false); this.loadRecords(); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.linking.set(false); },
    });
  }
}
