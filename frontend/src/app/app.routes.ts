import { Routes } from '@angular/router';
import { authGuard } from './guards/auth.guard';

export const routes: Routes = [
  {
    path: 'login',
    loadComponent: () => import('./pages/login/login.component').then(m => m.LoginComponent)
  },
  {
    path: '',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/dashboard/dashboard.component').then(m => m.DashboardComponent)
  },
  {
    path: 'sheet/:id',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/sheet-editor/sheet-editor.component').then(m => m.SheetEditorComponent)
  },
  {
    path: 'doc/:id',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/doc-editor/doc-editor.component').then(m => m.DocEditorComponent)
  },
  {
    path: 'slide/:id',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/slide-editor/slide-editor.component').then(m => m.SlideEditorComponent)
  },
  { path: '**', redirectTo: '' }
];