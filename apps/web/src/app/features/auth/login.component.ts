import { Component, inject, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { InputTextModule } from 'primeng/inputtext';
import { PasswordModule } from 'primeng/password';
import { ButtonModule } from 'primeng/button';
import { MessageModule } from 'primeng/message';
import { AuthService } from '../../core/services/auth.service';

@Component({
  standalone: true,
  imports: [FormsModule, InputTextModule, PasswordModule, ButtonModule, MessageModule],
  template: `
    <div class="min-h-screen flex items-center justify-center bg-gray-50">
      <div class="bg-white rounded-2xl shadow-lg p-8 w-full max-w-md">
        <div class="text-center mb-8">
          <div class="w-16 h-16 bg-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <i class="pi pi-microsoft text-white text-3xl"></i>
          </div>
          <h1 class="text-2xl font-bold text-gray-900">TestM365</h1>
          <p class="text-gray-500 mt-1">Portal de pruebas de integración M365</p>
        </div>

        @if (error()) {
          <p-message severity="error" [text]="error()!" styleClass="w-full mb-4" />
        }

        <div class="space-y-4">
          <div>
            <label class="text-sm font-medium text-gray-700">Email</label>
            <input
              pInputText
              type="email"
              [(ngModel)]="email"
              placeholder="admin@testm365.local"
              class="w-full mt-1"
            />
          </div>
          <div>
            <label class="text-sm font-medium text-gray-700">Contraseña</label>
            <p-password [(ngModel)]="password" [feedback]="false" [toggleMask]="true" styleClass="w-full mt-1" inputStyleClass="w-full" />
          </div>
          <p-button
            label="Iniciar sesión"
            icon="pi pi-sign-in"
            styleClass="w-full"
            [loading]="loading()"
            (onClick)="login()"
          />
        </div>

        <p class="text-center text-xs text-gray-400 mt-6">
          Admin: admin@testm365.local — Cliente: cliente@testm365.local<br>Contraseña: Admin1234!
        </p>
      </div>
    </div>
  `,
})
export class LoginComponent {
  private auth = inject(AuthService);
  private router = inject(Router);

  email = 'admin@testm365.local';
  password = 'Admin1234!';
  loading = signal(false);
  error = signal<string | null>(null);

  login() {
    this.loading.set(true);
    this.error.set(null);
    this.auth.login(this.email, this.password).subscribe({
      next: () => this.router.navigate(['/dashboard']),
      error: (e) => {
        this.error.set(e.error?.message ?? 'Error al iniciar sesión');
        this.loading.set(false);
      },
    });
  }
}
