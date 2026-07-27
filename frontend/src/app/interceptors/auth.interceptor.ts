import { HttpInterceptorFn, HttpErrorResponse } from '@angular/common/http';
import { inject } from '@angular/core';
import { Router } from '@angular/router';
import { catchError, throwError } from 'rxjs';

export const authInterceptor: HttpInterceptorFn = (req, next) => {
  const token = localStorage.getItem('auth_token');
  if (token) {
    req = req.clone({ setHeaders: { Authorization: `Bearer ${token}` } });
  }

  const router = inject(Router);

  return next(req).pipe(
    catchError((err: HttpErrorResponse) => {
      if (err.status === 401) {
        localStorage.removeItem('auth_token');
        localStorage.removeItem('auth_user');

        const hostname = window.location.hostname;
        let domainStr = '';
        if (hostname.includes('vsnaptechnology.com')) {
          domainStr = '; domain=.vsnaptechnology.com';
        }
        document.cookie = `auth_token=; path=/${domainStr}; expires=Thu, 01 Jan 1970 00:00:00 GMT; SameSite=None; Secure`;
        document.cookie = `auth_user=; path=/${domainStr}; expires=Thu, 01 Jan 1970 00:00:00 GMT; SameSite=None; Secure`;

        router.navigate(['/login']);
      }
      return throwError(() => err);
    })
  );
};
