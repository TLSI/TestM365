import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { TableModule } from 'primeng/table';
import { DialogModule } from 'primeng/dialog';
import { ToastModule } from 'primeng/toast';
import { MessageModule } from 'primeng/message';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

@Component({
  standalone: true,
  imports: [CommonModule, CardModule, ButtonModule, TableModule, DialogModule, ToastModule, MessageModule],
  providers: [MessageService],
  template: `
    <div class="p-6">
      <p-toast />
      <div class="flex items-center justify-between mb-6">
        <div>
          <h2 class="text-2xl font-bold text-gray-900">Mis Documentos</h2>
          <p class="text-gray-500 mt-1">Documentos gestionados de forma segura para su empresa</p>
        </div>
        <p-button label="Actualizar" icon="pi pi-refresh" severity="secondary" [loading]="loading()" (onClick)="load()" />
      </div>

      @if (!clientId()) {
        <p-message severity="info" text="No hay expediente asociado a esta cuenta. Contacte con su gestor." styleClass="w-full" />
      } @else {
        <p-card>
          <ng-template pTemplate="content">
            <p-table [value]="documents()" [loading]="loading()" styleClass="p-datatable-sm">
              <ng-template pTemplate="header">
                <tr>
                  <th>Nombre</th>
                  <th>Tipo</th>
                  <th>Tamaño</th>
                  <th>Acciones</th>
                </tr>
              </ng-template>
              <ng-template pTemplate="body" let-doc>
                <tr>
                  <td>
                    <div class="flex items-center gap-2">
                      <i [class]="getIcon(doc) + ' text-blue-500'"></i>
                      <span class="font-medium text-sm">{{ doc.name }}</span>
                    </div>
                  </td>
                  <td class="text-xs text-gray-500">{{ doc.file?.mimeType ?? doc.mimeType ?? 'Carpeta' }}</td>
                  <td class="text-xs text-gray-500">{{ doc.size ? (doc.size / 1024 | number:'1.0-0') + ' KB' : '—' }}</td>
                  <td>
                    <div class="flex gap-2">
                      @if (doc.file) {
                        <p-button label="Ver" size="small" [text]="true" icon="pi pi-eye" (onClick)="preview(doc)" />
                        <p-button label="Editar" size="small" severity="secondary" [text]="true" icon="pi pi-pencil" (onClick)="edit(doc)" />
                        <p-button label="Descargar" size="small" severity="success" [text]="true" icon="pi pi-download" (onClick)="download(doc)" />
                      }
                    </div>
                  </td>
                </tr>
              </ng-template>
              <ng-template pTemplate="emptymessage">
                <tr><td colspan="4" class="text-center text-gray-400 py-12">
                  <i class="pi pi-folder-open text-4xl mb-3 block text-gray-300"></i>
                  No hay documentos disponibles aún
                </td></tr>
              </ng-template>
            </p-table>
          </ng-template>
        </p-card>
      }

      <!-- Preview dialog -->
      <p-dialog header="Vista previa" [(visible)]="showPreview" [style]="{ width: '85vw', height: '80vh' }" [modal]="true">
        @if (previewUrl()) {
          <iframe [src]="previewUrl()!" class="w-full border-0" style="height: 70vh;" allow="fullscreen"></iframe>
        }
      </p-dialog>
    </div>
  `,
})
export class DocumentsComponent implements OnInit {
  private http = inject(HttpClient);
  private sanitizer = inject(DomSanitizer);
  private msg = inject(MessageService);

  clientId = signal<string | null>(null);
  documents = signal<unknown[]>([]);
  loading = signal(false);
  showPreview = false;
  previewUrl = signal<SafeResourceUrl | null>(null);

  ngOnInit() {
    this.http.get<{ data: { id: string }[] }>(`${environment.apiUrl}/clients`).subscribe({
      next: ({ data }) => {
        if (data.length > 0) {
          this.clientId.set(data[0].id);
          this.load();
        }
      },
    });
  }

  load() {
    const id = this.clientId();
    if (!id) return;
    this.loading.set(true);
    this.http.get<{ data: unknown[] }>(`${environment.apiUrl}/documents/clients/${id}`).subscribe({
      next: ({ data }) => { this.documents.set(data); this.loading.set(false); },
      error: () => this.loading.set(false),
    });
  }

  preview(doc: Record<string, unknown>) {
    this.http.get<{ data: { previewUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc['id']}/preview`
    ).subscribe({
      next: ({ data }) => {
        this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(data.previewUrl));
        this.showPreview = true;
      },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  edit(doc: Record<string, unknown>) {
    this.http.get<{ data: { editUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc['id']}/edit-link`
    ).subscribe({
      next: ({ data }) => window.open(data.editUrl, '_blank'),
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  download(doc: Record<string, unknown>) {
    window.open(
      `${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc['id']}/download`,
      '_blank'
    );
  }

  getIcon(doc: Record<string, unknown>): string {
    const mime = (doc['file'] as Record<string, unknown> | undefined)?.['mimeType'] as string ?? '';
    if (mime.includes('pdf')) return 'pi pi-file-pdf';
    if (mime.includes('word') || mime.includes('document')) return 'pi pi-file-word';
    if (mime.includes('excel') || mime.includes('spreadsheet')) return 'pi pi-file-excel';
    if (mime.includes('image')) return 'pi pi-image';
    if (!doc['file']) return 'pi pi-folder';
    return 'pi pi-file';
  }
}
