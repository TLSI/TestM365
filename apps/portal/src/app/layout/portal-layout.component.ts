import { Component, inject } from '@angular/core';
import { RouterOutlet, RouterLink, RouterLinkActive } from '@angular/router';
import { ButtonModule } from 'primeng/button';
import { AuthService } from '../core/services/auth.service';

@Component({
  standalone: true,
  imports: [RouterOutlet, RouterLink, RouterLinkActive, ButtonModule],
  template: `
    <div class="flex h-screen bg-gray-50">
      <!-- Sidebar -->
      <aside class="w-56 bg-white border-r border-gray-200 flex flex-col">
        <div class="p-5 border-b border-gray-100">
          <div class="flex items-center gap-2">
            <div class="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <i class="pi pi-building text-white text-sm"></i>
            </div>
            <div>
              <p class="font-bold text-gray-900 text-sm leading-none">Portal</p>
              <p class="text-xs text-gray-500">Cliente</p>
            </div>
          </div>
        </div>

        <nav class="flex-1 p-3 space-y-1">
          @for (item of navItems; track item.path) {
            <a
              [routerLink]="item.path"
              routerLinkActive="bg-blue-50 text-blue-700 font-medium"
              class="flex items-center gap-3 px-3 py-2 rounded-lg text-sm text-gray-600 hover:bg-gray-50 transition-colors"
            >
              <i [class]="item.icon + ' text-sm'"></i>
              {{ item.label }}
            </a>
          }
        </nav>

        <div class="p-3 border-t border-gray-100">
          <p-button
            label="Cerrar sesión"
            icon="pi pi-sign-out"
            severity="secondary"
            [text]="true"
            size="small"
            styleClass="w-full"
            (onClick)="logout()"
          />
        </div>
      </aside>

      <!-- Main content -->
      <main class="flex-1 overflow-auto">
        <router-outlet />
      </main>
    </div>
  `,
})
export class PortalLayoutComponent {
  private auth = inject(AuthService);

  navItems = [
    { path: '/dashboard', icon: 'pi pi-home', label: 'Inicio' },
    { path: '/documents', icon: 'pi pi-folder', label: 'Documentos' },
    { path: '/calendar', icon: 'pi pi-calendar', label: 'Calendario' },
  ];

  logout() { this.auth.logout(); }
}
