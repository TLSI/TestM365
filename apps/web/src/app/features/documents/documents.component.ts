import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TableModule } from 'primeng/table';
import { DialogModule } from 'primeng/dialog';
import { SelectModule } from 'primeng/select';
import { ToastModule } from 'primeng/toast';
import { FileUploadModule } from 'primeng/fileupload';
import { InputTextModule } from 'primeng/inputtext';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, TableModule, DialogModule,
    SelectModule, ToastModule, FileUploadModule, InputTextModule],
  providers: [MessageService],
  template: `
    <div class="p-6">
      <p-toast />
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Documentos / SharePoint</h2>
        <p class="text-gray-500 mt-1">FORLOPD como proxy — el cliente nunca ve SharePoint</p>
      </div>

      <!-- Client selector -->
      <div class="flex items-end gap-4 mb-6">
        <div class="flex-1 max-w-sm">
          <label class="text-sm font-medium text-gray-700">Cliente</label>
          <p-select
            [(ngModel)]="selectedClientId"
            [options]="clients()"
            optionLabel="name"
            optionValue="id"
            placeholder="Seleccionar cliente..."
            styleClass="w-full mt-1"
            (onChange)="onClientChange()"
          />
        </div>
        <p-button label="Crear site SharePoint" icon="pi pi-plus" severity="secondary"
          [disabled]="!selectedClientId" [loading]="creatingSite()" (onClick)="createSite()" />
        <p-button label="Cargar documentos" icon="pi pi-refresh"
          [disabled]="!selectedClientId" [loading]="loadingDocs()" (onClick)="loadDocs()" />
      </div>

      <!-- Docs table -->
      <p-card>
        <ng-template pTemplate="content">
          <!-- Upload -->
          @if (selectedClientId) {
            <div class="mb-4 flex items-center gap-3">
              <p-fileupload
                mode="basic"
                chooseLabel="Subir archivo a SharePoint"
                [auto]="true"
                [customUpload]="true"
                (uploadHandler)="uploadFile($event)"
              />
              <input pInputText [(ngModel)]="uploadFolder" placeholder="Carpeta (ej. LOPD/Contratos)" class="max-w-xs" />
            </div>
          }

          <p-table [value]="documents()" [loading]="loadingDocs()" styleClass="p-datatable-sm">
            <ng-template pTemplate="header">
              <tr>
                <th>Nombre</th>
                <th>Tamaño</th>
                <th>Tipo</th>
                <th>Acciones</th>
              </tr>
            </ng-template>
            <ng-template pTemplate="body" let-doc>
              <tr>
                <td class="font-medium">{{ doc.name }}</td>
                <td class="text-xs text-gray-500">{{ doc.size ? (doc.size / 1024 | number:'1.0-0') + ' KB' : '—' }}</td>
                <td class="text-xs">{{ doc.file?.mimeType ?? doc.mimeType ?? '—' }}</td>
                <td>
                  <div class="flex gap-2">
                    <p-button label="Preview" size="small" [text]="true" icon="pi pi-eye" (onClick)="previewDoc(doc)" />
                    <p-button label="Editar" size="small" severity="secondary" [text]="true" icon="pi pi-pencil" (onClick)="editDoc(doc)" />
                    <p-button label="Descargar" size="small" severity="success" [text]="true" icon="pi pi-download" (onClick)="downloadDoc(doc)" />
                  </div>
                </td>
              </tr>
            </ng-template>
            <ng-template pTemplate="emptymessage">
              <tr><td colspan="4" class="text-center text-gray-400 py-8">Selecciona un cliente y carga documentos</td></tr>
            </ng-template>
          </p-table>
        </ng-template>
      </p-card>

      <!-- Preview dialog -->
      <p-dialog header="Vista previa del documento" [(visible)]="showPreview" [style]="{ width: '85vw', height: '80vh' }" [modal]="true">
        @if (previewUrl()) {
          <iframe [src]="previewUrl()!" class="w-full h-full border-0" style="min-height: 60vh;" allow="fullscreen"></iframe>
        }
      </p-dialog>
    </div>
  `,
})
export class DocumentsComponent implements OnInit {
  private http = inject(HttpClient);
  private sanitizer = inject(DomSanitizer);
  private msg = inject(MessageService);

  clients = signal<{ id: string; name: string }[]>([]);
  documents = signal<unknown[]>([]);
  selectedClientId = '';
  uploadFolder = '';
  loadingDocs = signal(false);
  creatingSite = signal(false);
  showPreview = false;
  previewUrl = signal<SafeResourceUrl | null>(null);

  ngOnInit() {
    this.http.get<{ data: { id: string; name: string }[] }>(`${environment.apiUrl}/clients`).subscribe({
      next: ({ data }) => this.clients.set(data),
    });
  }

  onClientChange() {
    this.documents.set([]);
    if (this.selectedClientId) this.loadDocs();
  }

  loadDocs() {
    this.loadingDocs.set(true);
    this.http.get<{ data: unknown[] }>(`${environment.apiUrl}/documents/clients/${this.selectedClientId}`).subscribe({
      next: ({ data }) => { this.documents.set(data); this.loadingDocs.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.loadingDocs.set(false); },
    });
  }

  createSite() {
    this.creatingSite.set(true);
    this.http.post(`${environment.apiUrl}/documents/clients/${this.selectedClientId}/site`, {}).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Site creado en SharePoint' }); this.creatingSite.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.creatingSite.set(false); },
    });
  }

  uploadFile(event: { files: File[] }) {
    const file = event.files[0];
    const form = new FormData();
    form.append('file', file);
    const url = `${environment.apiUrl}/documents/clients/${this.selectedClientId}/upload${this.uploadFolder ? '?folder=' + this.uploadFolder : ''}`;
    this.http.post(url, form).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Archivo subido a SharePoint' }); this.loadDocs(); },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  previewDoc(doc: Record<string, unknown>) {
    this.http.get<{ data: { previewUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc['id']}/preview`
    ).subscribe({
      next: ({ data }) => {
        this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(data.previewUrl));
        this.showPreview = true;
      },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  editDoc(doc: Record<string, unknown>) {
    this.http.get<{ data: { editUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc['id']}/edit-link`
    ).subscribe({
      next: ({ data }) => window.open(data.editUrl, '_blank'),
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  downloadDoc(doc: Record<string, unknown>) {
    window.open(
      `${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc['id']}/download`,
      '_blank',
    );
  }
}
