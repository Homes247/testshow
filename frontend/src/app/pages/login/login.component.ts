import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { AuthService } from '../../services/auth.service';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="page">
      <div class="card">
        <div class="logo">🗂️ Office Suite</div>

        <div class="tabs">
          <button class="tab" [class.active]="mode === 'login'" (click)="mode = 'login'; error = ''">Sign In</button>
          <button class="tab" [class.active]="mode === 'register'" (click)="mode = 'register'; error = ''">Create Account</button>
        </div>

        <!-- LOGIN -->
        <div *ngIf="mode === 'login'" class="form">
          <div class="field">
            <label>Email</label>
            <input type="email" [(ngModel)]="email" placeholder="you@example.com" (keydown.enter)="submit()" />
          </div>
          <div class="field">
            <label>Password</label>
            <input type="password" [(ngModel)]="password" placeholder="••••••••" (keydown.enter)="submit()" />
          </div>
          <div class="error" *ngIf="error">{{ error }}</div>
          <button class="btn" [disabled]="loading" (click)="submit()">
            {{ loading ? 'Signing in…' : 'Sign In' }}
          </button>
        </div>

        <!-- REGISTER -->
        <div *ngIf="mode === 'register'" class="form">
          <div class="field">
            <label>Full Name</label>
            <input type="text" [(ngModel)]="name" placeholder="John Doe" (keydown.enter)="submit()" />
          </div>
          <div class="field">
            <label>Email</label>
            <input type="email" [(ngModel)]="email" placeholder="you@example.com" (keydown.enter)="submit()" />
          </div>
          <div class="field">
            <label>Password</label>
            <input type="password" [(ngModel)]="password" placeholder="Min 6 characters" (keydown.enter)="submit()" />
          </div>
          <div class="error" *ngIf="error">{{ error }}</div>
          <button class="btn" [disabled]="loading" (click)="submit()">
            {{ loading ? 'Creating account…' : 'Create Account' }}
          </button>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .page {
      min-height: 100vh; display: flex; align-items: center; justify-content: center;
      background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
    }
    .card {
      background: #fff; border-radius: 16px; padding: 40px 36px;
      width: 100%; max-width: 400px; box-shadow: 0 20px 60px rgba(0,0,0,.2);
    }
    .logo { font-size: 22px; font-weight: 700; text-align: center; margin-bottom: 28px; color: #111; }

    .tabs { display: flex; border-bottom: 2px solid #e5e7eb; margin-bottom: 28px; }
    .tab {
      flex: 1; padding: 10px; background: none; border: none; cursor: pointer;
      font-size: 14px; font-weight: 500; color: #6b7280; border-bottom: 2px solid transparent;
      margin-bottom: -2px; transition: all .15s;
    }
    .tab.active { color: #2563eb; border-bottom-color: #2563eb; }

    .form { display: flex; flex-direction: column; gap: 16px; }
    .field { display: flex; flex-direction: column; gap: 6px; }
    label { font-size: 13px; font-weight: 500; color: #374151; }
    input {
      padding: 10px 12px; border: 1px solid #e5e7eb; border-radius: 8px;
      font-size: 14px; outline: none; transition: border-color .15s;
    }
    input:focus { border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,.1); }

    .error { background: #fef2f2; color: #dc2626; padding: 10px 12px; border-radius: 8px; font-size: 13px; }

    .btn {
      padding: 11px; background: #2563eb; color: #fff; border: none; border-radius: 8px;
      cursor: pointer; font-size: 15px; font-weight: 600; margin-top: 4px; transition: background .15s;
    }
    .btn:hover:not([disabled]) { background: #1d4ed8; }
    .btn[disabled] { opacity: 0.6; cursor: not-allowed; }
  `]
})
export class LoginComponent {
  mode: 'login' | 'register' = 'login';
  name = '';
  email = '';
  password = '';
  error = '';
  loading = false;

  constructor(private auth: AuthService, private router: Router) {
    if (this.auth.isLoggedIn) this.router.navigate(['/']);
  }

  submit() {
    this.error = '';
    if (!this.email || !this.password) { this.error = 'Please fill in all fields.'; return; }
    if (this.mode === 'register' && !this.name) { this.error = 'Please enter your name.'; return; }
    if (this.password.length < 6) { this.error = 'Password must be at least 6 characters.'; return; }

    this.loading = true;
    const call = this.mode === 'login'
      ? this.auth.login(this.email, this.password)
      : this.auth.register(this.name, this.email, this.password);

    call.subscribe({
      next: () => this.router.navigate(['/']),
      error: (err) => {
        this.error = err?.error?.detail ?? 'Something went wrong. Try again.';
        this.loading = false;
      }
    });
  }
}
