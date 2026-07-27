import { Injectable, inject, signal, computed } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Router } from '@angular/router';
import { tap } from 'rxjs/operators';
import { environment } from '../../../environments/environment';

interface ClientInfo {
  id: string;
  company: string | null;
  name: string;
}

interface User {
  id: string;
  email: string;
  name: string;
  role: string;
  client?: ClientInfo | null;
}

@Injectable({ providedIn: 'root' })
export class AuthService {
  private http = inject(HttpClient);
  private router = inject(Router);

  private _user = signal<User | null>(null);
  private _token = signal<string | null>(localStorage.getItem('portal_token'));

  user = this._user.asReadonly();
  isLoggedIn = computed(() => !!this._token());
  token = this._token.asReadonly();

  constructor() {
    if (this._token()) {
      this.http.get<User>(`${environment.apiUrl}/auth/me`).pipe(
        tap(user => this._user.set(user))
      ).subscribe({ error: () => this.logout() });
    }
  }

  login(email: string, password: string) {
    return this.http.post<{ data: { accessToken: string; user: User } }>(
      `${environment.apiUrl}/auth/login`,
      { email, password }
    ).pipe(
      tap(({ data }) => {
        localStorage.setItem('portal_token', data.accessToken);
        this._token.set(data.accessToken);
        this._user.set(data.user);
      })
    );
  }

  logout() {
    localStorage.removeItem('portal_token');
    this._token.set(null);
    this._user.set(null);
    this.router.navigate(['/auth/login']);
  }

  getToken() { return this._token(); }
}
