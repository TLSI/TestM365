import { Component, inject, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { CardModule } from 'primeng/card';
import { InputTextModule } from 'primeng/inputtext';
import { PasswordModule } from 'primeng/password';
import { ButtonModule } from 'primeng/button';
import { MessageModule } from 'primeng/message';
import { AuthService } from '../../core/services/auth.service';

@Component({
  standalone: true,
  imports: [FormsModule, CardModule, InputTextModule, PasswordModule, ButtonModule, MessageModule],
  template: `
    <div class="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
      <div class="w-full max-w-sm">
        <div class="text-center mb-8">
          <div class="w-14 h-14 bg-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <i class="pi pi-building text-white text-2xl"></i>
          </div>
          <h1 class="text-2xl font-bold text-gray-900">Portal del Cliente</h1>
          <p class="text-gray-500 text-sm mt-1">Accede a tus documentos y citas</p>
        </div>

        <p-card>
          <ng-template pTemplate="content">
            @if (error()) {
              <p-message severity="error" [text]="error()!" styleClass="w-full mb-4" />
            }
            <div class="space-y-4">
              <div>
                <label class="text-sm font-medium text-gray-700">Email</label>
                <input pInputText [(ngModel)]="email" class="w-full mt-1" placeholder="cliente@empresa.com" />
              </div>
              <div>
                <label class="text-sm font-medium text-gray-700">Contraseña</label>
                <p-password [(ngModel)]="password" [feedback]="false" [toggleMask]="true" styleClass="w-full mt-1" inputStyleClass="w-full" />
              </div>
              <p-button
                label="Entrar"
                icon="pi pi-sign-in"
                styleClass="w-full"
                [loading]="loading()"
                (onClick)="login()"
              />
            </div>
          </ng-template>
        </p-card>

        <p class="text-center text-xs text-gray-400 mt-4">Plataforma TestM365 — Demo de integración M365</p>
      </div>
    </div>
  `,
})
export class LoginComponent {
  private auth = inject(AuthService);
  private router = inject(Router);

  email = 'cliente@testm365.local';
  password = 'Admin1234!';
  loading = signal(false);
  error = signal<string | null>(null);

  login() {
    this.loading.set(true);
    this.error.set(null);
    this.auth.login(this.email, this.password).subscribe({
      next: () => this.router.navigate(['/']),
      error: (e) => { this.error.set(e.error?.message ?? 'Credenciales incorrectas'); this.loading.set(false); },
    });
  }
}
