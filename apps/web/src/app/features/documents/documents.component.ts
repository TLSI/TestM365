import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { DialogModule } from 'primeng/dialog';
import { SelectModule } from 'primeng/select';
import { ToastModule } from 'primeng/toast';
import { FileUploadModule } from 'primeng/fileupload';
import { InputTextModule } from 'primeng/inputtext';
import { MessageService } from 'primeng/api';
import { environment } from '../../../environments/environment';

interface SpDocument {
  id: string;
  name: string;
  size?: number;
  file?: { mimeType: string };
  folder?: { childCount: number };
  mimeType?: string;
  lastModifiedDateTime?: string;
}

@Component({
  standalone: true,
  imports: [CommonModule, FormsModule, CardModule, ButtonModule, DialogModule,
    SelectModule, ToastModule, FileUploadModule, InputTextModule],
  providers: [MessageService],
  template: `
    <div class="p-6">
      <p-toast />
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900">Documentos / SharePoint</h2>
        <p class="text-gray-500 mt-1">FORLOPD como proxy — el cliente nunca ve SharePoint</p>
      </div>

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

      <p-card>
        <ng-template pTemplate="content">
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

          @if (loadingDocs()) {
            <div class="text-center py-8 text-gray-400"><i class="pi pi-spinner pi-spin text-2xl"></i></div>
          } @else if (documents().length === 0 && !selectedClientId) {
            <p class="text-center text-gray-400 py-8">Selecciona un cliente y carga documentos</p>
          } @else {
            <!-- Breadcrumb -->
            @if (currentPath().length > 0) {
              <nav class="flex items-center gap-1 text-sm mb-4 text-gray-500">
                <button class="hover:text-blue-600 font-medium" (click)="navigateToBreadcrumb(-1)">Raíz</button>
                @for (segment of currentPath(); track $index) {
                  <span>/</span>
                  <button class="hover:text-blue-600" (click)="navigateToBreadcrumb($index + 1)">{{ segment }}</button>
                }
              </nav>
            }
            <div class="overflow-x-auto">
              <table class="w-full text-sm">
                <thead>
                  <tr class="border-b border-gray-200 text-left text-xs text-gray-500 uppercase tracking-wide">
                    <th class="pb-3 pr-4 font-medium">Nombre</th>
                    <th class="pb-3 pr-4 font-medium">Tamaño</th>
                    <th class="pb-3 pr-4 font-medium">Tipo</th>
                    <th class="pb-3 font-medium">Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  @for (doc of documents(); track doc.id) {
                    <tr class="border-b border-gray-100 hover:bg-gray-50" [class.cursor-pointer]="doc.folder" (click)="doc.folder && navigateToFolder(doc)">
                      <td class="py-3 pr-4">
                        <div class="flex items-center gap-2">
                          <i [class]="getIcon(doc) + ' text-blue-500'"></i>
                          <span class="font-medium">{{ doc.name }}</span>
                          @if (doc.folder) { <i class="pi pi-angle-right text-gray-400 text-xs"></i> }
                        </div>
                      </td>
                      <td class="py-3 pr-4 text-xs text-gray-500">{{ doc.size ? (doc.size / 1024 | number:'1.0-0') + ' KB' : '—' }}</td>
                      <td class="py-3 pr-4 text-xs text-gray-500">{{ doc.file?.mimeType ?? doc.mimeType ?? (doc.folder ? 'Carpeta' : '—') }}</td>
                      <td class="py-3">
                        @if (doc.file) {
                          <div class="flex gap-1">
                            <p-button label="Preview" size="small" [text]="true" icon="pi pi-eye" (onClick)="previewDoc(doc); $event.stopPropagation()" />
                            <p-button label="Editar" size="small" severity="secondary" [text]="true" icon="pi pi-pencil" (onClick)="editDoc(doc); $event.stopPropagation()" />
                            <p-button label="Descargar" size="small" severity="success" [text]="true" icon="pi pi-download" (onClick)="downloadDoc(doc); $event.stopPropagation()" />
                          </div>
                        }
                      </td>
                    </tr>
                  }
                </tbody>
              </table>
            </div>
          }
        </ng-template>
      </p-card>

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
  documents = signal<SpDocument[]>([]);
  currentPath = signal<string[]>([]);
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

  onClientChange() { this.documents.set([]); this.currentPath.set([]); if (this.selectedClientId) this.loadDocs(); }

  loadDocs() {
    this.loadingDocs.set(true);
    const folder = this.currentPath().join('/');
    const url = `${environment.apiUrl}/documents/clients/${this.selectedClientId}${folder ? '?folder=' + encodeURIComponent(folder) : ''}`;
    this.http.get<{ data: SpDocument[] }>(url).subscribe({
      next: ({ data }) => { this.documents.set(data); this.loadingDocs.set(false); },
      error: (e) => { this.msg.add({ severity: 'error', detail: e.error?.message }); this.loadingDocs.set(false); },
    });
  }

  navigateToFolder(doc: SpDocument) {
    this.currentPath.update(p => [...p, doc.name]);
    this.loadDocs();
  }

  navigateToBreadcrumb(upToIndex: number) {
    this.currentPath.update(p => upToIndex < 0 ? [] : p.slice(0, upToIndex));
    this.loadDocs();
  }

  getIcon(doc: SpDocument): string {
    if (doc.folder) return 'pi pi-folder';
    const mime = doc.file?.mimeType ?? '';
    if (mime.includes('pdf')) return 'pi pi-file-pdf';
    if (mime.includes('word') || mime.includes('document')) return 'pi pi-file-word';
    if (mime.includes('excel') || mime.includes('spreadsheet')) return 'pi pi-file-excel';
    if (mime.includes('image')) return 'pi pi-image';
    return 'pi pi-file';
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
    const finalFolder = [...this.currentPath(), ...this.uploadFolder ? [this.uploadFolder] : []].join('/');
    const url = `${environment.apiUrl}/documents/clients/${this.selectedClientId}/upload${finalFolder ? '?folder=' + encodeURIComponent(finalFolder) : ''}`;
    this.http.post(url, form).subscribe({
      next: () => { this.msg.add({ severity: 'success', summary: 'Archivo subido a SharePoint' }); this.loadDocs(); },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  previewDoc(doc: SpDocument) {
    this.http.get<{ data: { previewUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc.id}/preview`
    ).subscribe({
      next: ({ data }) => { this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(data.previewUrl)); this.showPreview = true; },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  editDoc(doc: SpDocument) {
    this.http.get<{ data: { editUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc.id}/edit-link`
    ).subscribe({
      next: ({ data }) => window.open(data.editUrl, '_blank'),
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  downloadDoc(doc: SpDocument) {
    window.open(`${environment.apiUrl}/documents/clients/${this.selectedClientId}/items/${doc.id}/download`, '_blank');
  }
}
