import { Component, inject } from '@angular/core';
import { RouterOutlet, RouterLink, RouterLinkActive } from '@angular/router';
import { CommonModule } from '@angular/common';
import { ButtonModule } from 'primeng/button';
import { AvatarModule } from 'primeng/avatar';
import { BadgeModule } from 'primeng/badge';
import { AuthService } from '../core/services/auth.service';

interface NavItem {
  label: string;
  icon: string;
  route: string;
}

@Component({
  selector: 'app-layout',
  standalone: true,
  imports: [RouterOutlet, RouterLink, RouterLinkActive, CommonModule, ButtonModule, AvatarModule, BadgeModule],
  template: `
    <div class="flex h-screen bg-gray-50 dark:bg-gray-900">
      <!-- Sidebar -->
      <aside class="w-64 bg-white dark:bg-gray-800 border-r border-gray-200 dark:border-gray-700 flex flex-col">
        <!-- Logo -->
        <div class="p-6 border-b border-gray-200 dark:border-gray-700">
          <div class="flex items-center gap-3">
            <div class="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <i class="pi pi-microsoft text-white text-sm"></i>
            </div>
            <div>
              <h1 class="font-bold text-gray-900 dark:text-white text-sm">TestM365</h1>
              <p class="text-xs text-gray-500">Admin Portal</p>
            </div>
          </div>
        </div>

        <!-- Navigation -->
        <nav class="flex-1 p-4 space-y-1 overflow-y-auto">
          @for (item of navItems; track item.route) {
            <a
              [routerLink]="item.route"
              routerLinkActive="bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300"
              class="flex items-center gap-3 px-3 py-2 rounded-lg text-gray-600 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors text-sm font-medium"
            >
              <i [class]="'pi ' + item.icon + ' text-base'"></i>
              {{ item.label }}
            </a>
          }
        </nav>

        <!-- User -->
        <div class="p-4 border-t border-gray-200 dark:border-gray-700">
          <div class="flex items-center gap-3 mb-3">
            <p-avatar [label]="userInitial()" shape="circle" styleClass="bg-blue-600 text-white" />
            <div class="flex-1 min-w-0">
              <p class="text-sm font-medium text-gray-900 dark:text-white truncate">{{ auth.user()?.name }}</p>
              <p class="text-xs text-gray-500 truncate">{{ auth.user()?.email }}</p>
            </div>
          </div>
          <p-button
            label="Cerrar sesión"
            icon="pi pi-sign-out"
            severity="secondary"
            size="small"
            [text]="true"
            styleClass="w-full"
            (onClick)="auth.logout()"
          />
        </div>
      </aside>

      <!-- Main content -->
      <main class="flex-1 overflow-y-auto">
        <router-outlet />
      </main>
    </div>
  `,
})
export class AppLayoutComponent {
  readonly auth = inject(AuthService);

  readonly navItems: NavItem[] = [
    { label: 'Dashboard', icon: 'pi-home', route: '/dashboard' },
    { label: 'Microsoft 365', icon: 'pi-microsoft', route: '/microsoft' },
    { label: 'Calendario', icon: 'pi-calendar', route: '/calendar' },
    { label: 'Correo', icon: 'pi-envelope', route: '/mail' },
    { label: 'Documentos', icon: 'pi-file', route: '/documents' },
    { label: 'Teams', icon: 'pi-comments', route: '/teams' },
  ];

  userInitial() {
    return this.auth.user()?.name?.charAt(0)?.toUpperCase() ?? 'A';
  }
}
