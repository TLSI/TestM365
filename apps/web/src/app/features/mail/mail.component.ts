import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TableModule } from 'primeng/table';
import { DialogModule } from 'primeng/dialog';
import { InputTextModule } from 'primeng/inputtext';
import { TextareaModule } from 'primeng/textarea';
import { TabsModule } from 'primeng/tabs';
import { TagModule } from 'primeng/tag';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, TableModule, DialogModule,
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
          <!-- Inbox -->
          <p-tabpanel value="0">
            <div class="flex justify-end mb-3">
              <p-button label="Cargar bandeja de Outlook" icon="pi pi-refresh" severity="secondary" [loading]="loadingInbox()" (onClick)="loadInbox()" />
            </div>
            <p-table [value]="inbox()" [loading]="loadingInbox()" styleClass="p-datatable-sm">
              <ng-template pTemplate="header">
                <tr>
                  <th>De</th>
                  <th>Asunto</th>
                  <th>Vista previa</th>
                  <th>Fecha</th>
                  <th>Vincular</th>
                </tr>
              </ng-template>
              <ng-template pTemplate="body" let-msg>
                <tr>
                  <td class="text-xs">{{ msg.from?.emailAddress?.address }}</td>
                  <td class="text-sm font-medium">{{ msg.subject }}</td>
                  <td class="text-xs text-gray-400 max-w-xs truncate">{{ msg.bodyPreview }}</td>
                  <td class="text-xs">{{ msg.receivedDateTime | date:'short' }}</td>
                  <td>
                    <p-button label="Vincular" size="small" [text]="true" (onClick)="openLinkDialog(msg)" />
                  </td>
                </tr>
              </ng-template>
              <ng-template pTemplate="emptymessage">
                <tr><td colspan="5" class="text-center text-gray-400 py-8">Pulsa "Cargar bandeja" para ver los emails de Outlook</td></tr>
              </ng-template>
            </p-table>
          </p-tabpanel>

          <!-- Send -->
          <p-tabpanel value="1">
            <div class="max-w-xl space-y-4">
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

          <!-- Linked -->
          <p-tabpanel value="2">
            <p-table [value]="records()" [loading]="loadingRecords()" styleClass="p-datatable-sm">
              <ng-template pTemplate="header">
                <tr><th>Dirección</th><th>De/Para</th><th>Asunto</th><th>Fecha</th><th>Cliente</th></tr>
              </ng-template>
              <ng-template pTemplate="body" let-r>
                <tr>
                  <td><p-tag [value]="r.direction" [severity]="r.direction === 'OUTBOUND' ? 'success' : 'info'" /></td>
                  <td class="text-xs">{{ r.direction === 'OUTBOUND' ? r.to?.join(', ') : r.from }}</td>
                  <td class="text-sm">{{ r.subject }}</td>
                  <td class="text-xs">{{ r.sentAt | date:'short' }}</td>
                  <td class="text-xs">{{ r.clientId ?? '—' }}</td>
                </tr>
              </ng-template>
            </p-table>
          </p-tabpanel>
        </p-tabpanels>
      </p-tabs>

      <!-- Link dialog -->
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

  inbox = signal<unknown[]>([]);
  records = signal<unknown[]>([]);
  loadingInbox = signal(false);
  loadingRecords = signal(false);
  sending = signal(false);
  linking = signal(false);
  showLinkDialog = false;
  linkClientId = '';
  selectedMsg = signal<Record<string, unknown> | null>(null);

  sendForm = { to: '', cc: '', subject: 'Prueba desde TestM365', body: '<p>Hola, este es un email de prueba enviado via Microsoft Graph API desde la plataforma TestM365.</p>' };

  ngOnInit() { this.loadRecords(); }

  loadInbox() {
    this.loadingInbox.set(true);
    this.http.get<{ data: unknown[] }>(`${environment.apiUrl}/mail/inbox`).subscribe({
      next: ({ data }) => { this.inbox.set(data); this.loadingInbox.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', summary: 'Error', detail: e.error?.message }); this.loadingInbox.set(false); },
    });
  }

  loadRecords() {
    this.loadingRecords.set(true);
    this.http.get<{ data: unknown[] }>(`${environment.apiUrl}/mail/records`).subscribe({
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
      next: () => { this.msg.add({ severity: 'success', summary: 'Email enviado', detail: 'Enviado via Outlook / Graph API' }); this.sending.set(false); this.loadRecords(); },
      error: (e) => { this.msg.add({ severity: 'error', summary: 'Error', detail: e.error?.message }); this.sending.set(false); },
    });
  }

  openLinkDialog(msg: Record<string, unknown>) {
    this.selectedMsg.set(msg);
    this.showLinkDialog = true;
  }

  linkEmail() {
    const id = (this.selectedMsg() as { id?: string })?.id;
    if (!id) return;
    this.linking.set(true);
    this.http.post(`${environment.apiUrl}/mail/link/${id}?clientId=${this.linkClientId}`, {}).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Email vinculado al cliente' }); this.showLinkDialog = false; this.linking.set(false); this.loadRecords(); },
      error: (e) => { this.msg.add({ severity: 'error', summary: 'Error', detail: e.error?.message }); this.linking.set(false); },
    });
  }
}
