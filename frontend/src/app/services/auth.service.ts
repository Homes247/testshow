import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable, tap, BehaviorSubject } from 'rxjs';
import { Router } from '@angular/router';
import { environment } from '../../environments/environment';

export interface AuthUser { id: string; name: string; email: string; phone?: string; avatar_url?: string; avatar_color?: string; }

@Injectable({ providedIn: 'root' })
export class AuthService {
  private base = (environment as any).authUrl || `${environment.apiUrl}/auth`;
  private _user: AuthUser | null = null;
  private _ready$ = new BehaviorSubject<boolean>(false);
  ready$ = this._ready$.asObservable();

  constructor(private http: HttpClient, private router: Router) {
    this.syncFromStorage();
  }

  private syncFromStorage() {
    const cookieToken = this.getCookie('auth_token');
    const cookieUser = this.getCookie('auth_user');

    if (cookieToken) {
      localStorage.setItem('auth_token', cookieToken);
    } else {
      localStorage.removeItem('auth_token');
    }

    if (cookieUser) {
      localStorage.setItem('auth_user', cookieUser);
    } else {
      localStorage.removeItem('auth_user');
    }

    const storedToken = localStorage.getItem('auth_token');
    const storedUser = localStorage.getItem('auth_user');

    if (storedToken) {
      if (storedUser) {
        try {
          // storedUser from localStorage is plain JSON; from cookie it may be URI-encoded
          const decoded = storedUser.startsWith('%') ? decodeURIComponent(storedUser) : storedUser;
          const parsed = JSON.parse(decoded);
          if (parsed && parsed.id != null) {
            parsed.id = String(parsed.id);
            this._user = parsed;
            this._ready$.next(true);  // fires synchronously — dashboard loads immediately
          }
        } catch { }
      }
      // Fetch fresh in background — does NOT block the dashboard load
      this.fetchMe();
    } else {
      this._ready$.next(true);
    }
  }

  private fetchMe() {
    this.http.get<AuthUser>(`${this.base}/me`, {
      headers: { Authorization: `Bearer ${localStorage.getItem('auth_token')}` }
    }).subscribe({
      next: (user: AuthUser) => {
        user.id = String(user.id);
        this._user = user;
        const userStr = JSON.stringify(user);
        localStorage.setItem('auth_user', userStr);
        this.setCookie('auth_user', encodeURIComponent(userStr));
        // Only emit ready if not already emitted (i.e. storedUser parse failed earlier)
        if (!this._ready$.value) this._ready$.next(true);
      },
      error: (err: any) => {
        if (err?.status === 401) {
          this.logout();
        }
      }
    });
  }

  get token(): string | null { 
    return localStorage.getItem('auth_token'); 
  }
  
  get user(): AuthUser | null { return this._user; }
  
  get isLoggedIn(): boolean { return !!this.token; }

  register(name: string, email: string, password: string): Observable<any> {
    return this.http.post(`${this.base}/register`, { name, email, password }).pipe(
      tap((res: any) => this.store(res))
    );
  }

  login(email: string, password: string): Observable<any> {
    return this.http.post(`${this.base}/login`, { email, password }).pipe(
      tap((res: any) => this.store(res))
    );
  }

  logout() {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('auth_user');
    this.deleteCookie('auth_token');
    this.deleteCookie('auth_user');
    this._user = null;
    this.router.navigate(['/login']);
  }

  private store(res: any) {
    if (res.user && res.user.id != null) {
      res.user.id = String(res.user.id);
    }
    const token = res.token || res.access_token;
    const userStr = JSON.stringify(res.user);
    
    if (token) {
      localStorage.setItem('auth_token', token);
      this.setCookie('auth_token', token);
    }
    
    if (res.user) {
      localStorage.setItem('auth_user', userStr);
      this.setCookie('auth_user', encodeURIComponent(userStr));
      this._user = res.user;
    }
    
    this._ready$.next(true);
  }

  private getCookie(name: string): string | null {
    const value = `; ${document.cookie}`;
    const parts = value.split(`; ${name}=`);
    if (parts.length === 2) return parts.pop()?.split(';').shift() || null;
    return null;
  }

  private setCookie(name: string, value: string) {
    const hostname = window.location.hostname;
    let domainStr = '';
    if (hostname.includes('vsnaptechnology.com')) {
      domainStr = '; domain=.vsnaptechnology.com';
    }
    document.cookie = `${name}=${value}; path=/${domainStr}; max-age=2592000; SameSite=None; Secure`; // 30 days
  }

  private deleteCookie(name: string) {
    const hostname = window.location.hostname;
    let domainStr = '';
    if (hostname.includes('vsnaptechnology.com')) {
      domainStr = '; domain=.vsnaptechnology.com';
    }
    document.cookie = `${name}=; path=/${domainStr}; expires=Thu, 01 Jan 1970 00:00:00 GMT; SameSite=None; Secure`;
  }
}
