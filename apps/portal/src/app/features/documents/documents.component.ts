import { Component, OnInit, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { CardModule } from 'primeng/card';
import { ButtonModule } from 'primeng/button';
import { DialogModule } from 'primeng/dialog';
import { ToastModule } from 'primeng/toast';
import { MessageModule } from 'primeng/message';
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
  imports: [CommonModule, CardModule, ButtonModule, DialogModule, ToastModule, MessageModule],
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
            @if (loading()) {
              <div class="text-center py-8 text-gray-400"><i class="pi pi-spinner pi-spin text-2xl"></i></div>
            } @else if (documents().length === 0) {
              <div class="text-center py-12 text-gray-400">
                <i class="pi pi-folder-open text-4xl mb-3 block text-gray-300"></i>
                <p>No hay documentos disponibles aún</p>
              </div>
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
                      <th class="pb-3 pr-4 font-medium">Tipo</th>
                      <th class="pb-3 pr-4 font-medium">Tamaño</th>
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
                        <td class="py-3 pr-4 text-xs text-gray-500">{{ doc.file?.mimeType ?? doc.mimeType ?? (doc.folder ? 'Carpeta' : '—') }}</td>
                        <td class="py-3 pr-4 text-xs text-gray-500">{{ doc.size ? (doc.size / 1024 | number:'1.0-0') + ' KB' : '—' }}</td>
                        <td class="py-3">
                          @if (doc.file) {
                            <div class="flex gap-1">
                              <p-button label="Ver" size="small" [text]="true" icon="pi pi-eye" (onClick)="preview(doc); $event.stopPropagation()" />
                              <p-button label="Editar" size="small" severity="secondary" [text]="true" icon="pi pi-pencil" (onClick)="edit(doc); $event.stopPropagation()" />
                              <p-button label="Descargar" size="small" severity="success" [text]="true" icon="pi pi-download" (onClick)="download(doc); $event.stopPropagation()" />
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
      }

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
  documents = signal<SpDocument[]>([]);
  currentPath = signal<string[]>([]);
  loading = signal(false);
  showPreview = false;
  previewUrl = signal<SafeResourceUrl | null>(null);

  ngOnInit() {
    this.http.get<{ data: { id: string }[] }>(`${environment.apiUrl}/clients`).subscribe({
      next: ({ data }) => { if (data.length > 0) { this.clientId.set(data[0].id); this.load(); } },
    });
  }

  load() {
    const id = this.clientId();
    if (!id) return;
    this.loading.set(true);
    const folder = this.currentPath().join('/');
    const url = `${environment.apiUrl}/documents/clients/${id}${folder ? '?folder=' + encodeURIComponent(folder) : ''}`;
    this.http.get<{ data: SpDocument[] }>(url).subscribe({
      next: ({ data }) => { this.documents.set(data); this.loading.set(false); },
      error: () => this.loading.set(false),
    });
  }

  navigateToFolder(doc: SpDocument) {
    this.currentPath.update(p => [...p, doc.name]);
    this.load();
  }

  navigateToBreadcrumb(upToIndex: number) {
    this.currentPath.update(p => upToIndex < 0 ? [] : p.slice(0, upToIndex));
    this.load();
  }

  preview(doc: SpDocument) {
    this.http.get<{ data: { previewUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc.id}/preview`
    ).subscribe({
      next: ({ data }) => { this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(data.previewUrl)); this.showPreview = true; },
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  edit(doc: SpDocument) {
    this.http.get<{ data: { editUrl: string } }>(
      `${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc.id}/edit-link`
    ).subscribe({
      next: ({ data }) => window.open(data.editUrl, '_blank'),
      error: (e) => this.msg.add({ severity: 'error', detail: e.error?.message }),
    });
  }

  download(doc: SpDocument) {
    window.open(`${environment.apiUrl}/documents/clients/${this.clientId()}/items/${doc.id}/download`, '_blank');
  }

  getIcon(doc: SpDocument): string {
    const mime = doc.file?.mimeType ?? '';
    if (mime.includes('pdf')) return 'pi pi-file-pdf';
    if (mime.includes('word') || mime.includes('document')) return 'pi pi-file-word';
    if (mime.includes('excel') || mime.includes('spreadsheet')) return 'pi pi-file-excel';
    if (mime.includes('image')) return 'pi pi-image';
    if (!doc.file) return 'pi pi-folder';
    return 'pi pi-file';
  }
}
