import { Component, OnInit, OnDestroy, HostListener, ViewChild, ElementRef, Pipe, PipeTransform } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { ChatWidgetComponent } from '../../components/chat-widget/chat-widget.component';
import { ApiService } from '../../services/api.service';
import { AuthService } from '../../services/auth.service';

@Pipe({
  name: 'safeHtml',
  standalone: true
})
export class SafeHtmlPipe implements PipeTransform {
  constructor(private sanitizer: DomSanitizer) {}
  transform(value: string): SafeHtml {
    return this.sanitizer.bypassSecurityTrustHtml(value || '');
  }
}

@Component({
  selector: 'app-doc-editor',
  standalone: true,
  imports: [CommonModule, FormsModule, ChatWidgetComponent, SafeHtmlPipe],
  template: `
    <div class="shell" (click)="closeMenus($event)" [class.theme-dark]="isDarkMode" [class.reader-mode]="isReaderView">
      <!-- Header Bar -->
      <div class="header-bar">
        <div class="header-left">
          <button class="back-btn" (click)="back()" title="Back to Dashboard" style="background:none; border:none; cursor:pointer; color:inherit; display:flex; align-items:center; justify-content:center; width:32px; height:32px; border-radius:50%; flex-shrink:0; opacity:0.8; margin-right:4px;">
            <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10 19l-7-7m0 0l7-7m-7 7h18"/></svg>
          </button>
          <div class="brand-btn" title="Writer" style="cursor:pointer;" (click)="goHome()">
            <span class="material-symbols-outlined" style="font-size:28px;">description</span>
            <span class="brand-label">Writer</span>
          </div>
          <div class="doc-meta">
            <div class="title-row">
              <input class="title-input" [(ngModel)]="title" (blur)="save()" placeholder="Untitled Document" [style.width.ch]="(title || 'Untitled Document').length + 3" />
              <button class="star-btn" (click)="toggleStar()" title="Star">
                <span class="material-symbols-outlined" [class.filled]="isStarred">{{ isStarred ? 'star' : 'star_border' }}</span>
              </button>
              <span class="save-status">
                <span class="material-symbols-outlined" style="font-size:15px;">cloud_done</span> 
                {{ isSaving ? 'Saving...' : savedTimeStr }}
              </span>
            </div>
            
          </div>
        </div>

        <div class="header-right">
          <span class="collab-badge" *ngIf="activeUsers > 1">
            <span class="material-symbols-outlined" style="font-size:16px;">group</span>
            {{ activeUsers }}
          </span>
          <button class="share-btn" (click)="shareModalOpen = true; closeMenus()">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
            Share
          </button>
          <div style="position:relative;">
            <button class="header-icon-btn" (click)="toggleMenu('notif', $event)" title="Notifications"><span class="material-symbols-outlined">notifications</span></button>
            <div class="dropdown notif-dd shadow-lg" *ngIf="activeMenu === 'notif'" (click)="$event.stopPropagation()">
              <div style="padding:12px; font-weight:600; border-bottom:1px solid #e0e3e8; font-size:12px; color:#5f6368;">NOTIFICATION</div>
              <div style="padding:24px 16px; text-align:center; color:#5f6368; font-size:13px;">You haven't received any notifications yet</div>
            </div>
          </div>
          
          <div style="position:relative;">
            <div class="avatar" (click)="toggleMenu('profile', $event)" [title]="auth.user?.name ?? ''" style="cursor:pointer;">{{ initials }}</div>
            <div class="dropdown profile-dd shadow-lg" *ngIf="activeMenu === 'profile'" (click)="$event.stopPropagation()">
              <div style="padding:16px; text-align:center; border-bottom:1px solid #e0e3e8;">
                <div style="width:48px; height:48px; border-radius:50%; background:linear-gradient(135deg, #667eea, #764ba2); color:#fff; display:flex; align-items:center; justify-content:center; font-size:20px; font-weight:600; margin:0 auto 8px;">{{ initials }}</div>
                <div style="font-weight:600; font-size:14px; color:#202124;">{{ auth.user?.email || 'admin@vsnapmail.co.in' }}</div>
                <div style="font-size:12px; color:#5f6368; margin-top:2px;">User ID: {{ auth.user?.id || 1 }}</div>
              </div>
              <div style="padding:8px 0;">
                <div class="dd-item" (click)="auth.logout()">
                  <span class="material-symbols-outlined dd-icon">logout</span>
                  <span class="dd-text">Sign Out</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="menu-bar-row" style="display:flex; align-items:center; padding: 2px 12px; background: #fff; border-bottom: 1px solid #dadce0; position: relative; z-index: 205;">
            <div class="menu-bar" (mousedown)="$event.preventDefault()">
              <div class="menu-item" (click)="toggleMenu('file', $event)" [class.active]="activeMenu === 'file'">
                File
                <div class="dropdown file-dd" *ngIf="activeMenu === 'file'" (scroll)="closeAllSubmenus()">
                  <!-- New Document -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, sub1)">
                     <span class="material-symbols-outlined dd-icon">note_add</span><span class="dd-text">New Document</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                     <div #sub1 class="sub-dropdown">
                       <div class="dd-item" (click)="newDoc()"><span class="material-symbols-outlined dd-icon">draft</span><span class="dd-text">Blank Document</span></div>
                       <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, sub12)">
                          <span class="material-symbols-outlined dd-icon">post_add</span><span class="dd-text">Document Using Template...</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                          <div #sub12 class="sub-dropdown">
                             <div class="dd-item" (click)="loadTemplate('resume')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Resume</span></div>
                             <div class="dd-item" (click)="loadTemplate('letter')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Letter</span></div>
                             <div class="dd-item" (click)="loadTemplate('proposal')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Project Proposal</span></div>
                             <div class="dd-item" (click)="loadTemplate('notes')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Meeting Notes</span></div>
                             <div class="dd-item" (click)="loadTemplate('brochure')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Brochure</span></div>
                             <div class="dd-item" (click)="loadTemplate('newsletter')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Newsletter</span></div>
                             <div class="dd-item" (click)="loadTemplate('invoice')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Invoice</span></div>
                             <div class="dd-item" (click)="loadTemplate('report')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Business Report</span></div>
                             <div class="dd-item" (click)="loadTemplate('plan')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Business Plan</span></div>
                             <div class="dd-item" (click)="loadTemplate('essay')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span><span class="dd-text">Essay / Academic</span></div>
                             <div class="dd-sep"></div>
                             <div class="dd-item" (click)="showToast('Template Gallery opening')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">grid_view</span><span class="dd-text">Template Gallery...</span></div>
                          </div>
                       </div>
                     </div>
                  </div>

                  <!-- New Automation Template -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subAutoTemp)">
                    <span class="material-symbols-outlined dd-icon">auto_awesome_mosaic</span><span class="dd-text">New Automation Template</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                    <div #subAutoTemp class="sub-dropdown">
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMerge)">
                        <span class="material-symbols-outlined dd-icon" style="margin-top:2px;">merge_type</span>
                        <div style="display:flex; flex-direction:column; flex:1;">
                          <span class="dd-text" style="color:#1a73e8;">Merge Template</span>
                          <span style="font-size:10px; color:#5f6368; line-height:1;">Generate personalized documents in bulk</span>
                        </div>
                        <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                        <div #subMerge class="sub-dropdown">
                           <div class="dd-item" (click)="showToast('Design From Scratch')" style="align-items:flex-start;">
                             <div style="display:flex; flex-direction:column;">
                               <span class="dd-text" style="color:#1a73e8;">Design From Scratch</span>
                               <span style="font-size:10px; color:#5f6368; line-height:1; margin-top:2px;">Build custom templates using a blank canvas</span>
                             </div>
                           </div>
                           <div class="dd-item" (click)="showToast('Design Over PDF')" style="align-items:flex-start;">
                             <div style="display:flex; flex-direction:column;">
                               <span class="dd-text">Design Over PDF</span>
                               <span style="font-size:10px; color:#5f6368; line-height:1; margin-top:2px;">Map merge fields to an existing PDF</span>
                             </div>
                           </div>
                        </div>
                      </div>
                      <div class="dd-item" (click)="showToast('Fillable Template')" style="align-items:flex-start;">
                        <span class="material-symbols-outlined dd-icon" style="margin-top:2px;">data_object</span>
                        <div style="display:flex; flex-direction:column; flex:1;">
                          <span class="dd-text">Fillable Template</span>
                          <span style="font-size:10px; color:#5f6368; line-height:1;">Create and publish forms for users to fill</span>
                        </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSignTemp)" style="align-items:flex-start;">
                        <span class="material-symbols-outlined dd-icon" style="margin-top:2px; color:#0f9d58;">draw</span>
                        <div style="display:flex; flex-direction:column; flex:1;">
                          <span class="dd-text">Sign Template</span>
                          <span style="font-size:10px; color:#5f6368; line-height:1;">Collect signatures by automating document delivery</span>
                        </div>
                        <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                        <div #subSignTemp class="sub-dropdown">
                           <div class="dd-item" (click)="showToast('Sign with VMail Sign')"><span class="dd-text">Sign with VMail Sign</span></div>
                        </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subLabelTemp)" style="align-items:flex-start;">
                        <span class="material-symbols-outlined dd-icon" style="margin-top:2px;">label</span>
                        <div style="display:flex; flex-direction:column; flex:1;">
                          <span class="dd-text">Label Template</span>
                          <span style="font-size:10px; color:#5f6368; line-height:1;">Create formatted labels ready for printing</span>
                        </div>
                        <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                        <div #subLabelTemp class="sub-dropdown">
                           <div class="dd-item" (click)="showToast('Label Template')"><span class="dd-text">Create Label</span></div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <!-- Import -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subImport)">
                    <span class="material-symbols-outlined dd-icon">login</span><span class="dd-text">Import</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                    <div #subImport class="sub-dropdown">
                       <div class="dd-item" (click)="triggerImport($event)"><span class="material-symbols-outlined dd-icon" style="font-size:16px; color:#1a73e8;">computer</span><span class="dd-text" style="color:#1a73e8;">From Computer</span></div>
                       <div class="dd-item" (click)="cloudModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">cloud_download</span><span class="dd-text">From Cloud Drives</span></div>
                    </div>
                  </div>

                  <!-- Open -->
                  <div class="dd-item has-sub" (click)="docInput.click(); closeMenus()" (mouseenter)="positionSubmenu($event, subOpen)">
                    <span class="material-symbols-outlined dd-icon">folder_open</span><span class="dd-text">Open</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                    <div #subOpen class="sub-dropdown">
                       <div style="padding: 4px 12px; font-size: 11px; font-weight: 600; color: #202124;">Recently Opened</div>
                       <div class="dd-sep"></div>
                       <ng-container *ngIf="recentDocs.length > 0">
                         <div class="dd-item" *ngFor="let d of recentDocs" (click)="openDoc(d.id)">
                           <span class="material-symbols-outlined dd-icon" style="font-size:16px;">description</span>
                           <span class="dd-text" [title]="d.title">{{ (d.title && d.title.length > 30) ? (d.title | slice:0:30) + '...' : d.title || 'Untitled Document' }}</span>
                         </div>
                       </ng-container>
                       <div *ngIf="recentDocs.length === 0" style="padding: 4px 12px; font-size: 12px; color: #5f6368;">No recent documents</div>
                       <div class="dd-sep"></div>
                       <div class="dd-item" (click)="docInput.click(); closeMenus()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">folder</span><span class="dd-text">More Documents...</span></div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <!-- Manage Documents -->
                  <div class="dd-item" (click)="showToast('Manage Documents')"><span class="material-symbols-outlined dd-icon">folder_managed</span><span class="dd-text">Manage Documents</span></div>

                  <div class="dd-sep"></div>

                  <!-- Make a Copy -->
                  <div class="dd-item" (click)="makeCopy()"><span class="material-symbols-outlined dd-icon">content_copy</span><span class="dd-text">Make a Copy...</span><span class="dd-hint">Ctrl+Shift+S</span></div>
                  
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSaveAs)">
                    <span class="material-symbols-outlined dd-icon">save</span><span class="dd-text">Save As</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                    <div #subSaveAs class="sub-dropdown">
                       <div class="dd-item" (click)="saveAs()"><span class="material-symbols-outlined dd-icon" style="font-size:16px; color:#1a73e8;">save</span><span class="dd-text" style="color:#1a73e8;">Save As...</span></div>
                       <div class="dd-item" (click)="saveAsTemplate()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">receipt_long</span><span class="dd-text">Save As Template...</span></div>
                       <div class="dd-item" (click)="cloudModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">cloud_upload</span><span class="dd-text">Save to Other Cloud Drives</span></div>
                    </div>
                  </div>

                  <!-- AutoSave Toggle -->
                  <div class="dd-item" (click)="toggleAutoSave(); $event.stopPropagation()">
                    <span class="material-symbols-outlined dd-icon" [style.color]="autoSaveEnabled ? '#1a73e8' : '#5f6368'">{{ autoSaveEnabled ? 'check_box' : 'check_box_outline_blank' }}</span>
                    <span class="dd-text">AutoSave</span>
                  </div>

                  <!-- Manual Save (only if autosave is disabled) -->
                  <div class="dd-item" *ngIf="!autoSaveEnabled" (click)="save(true); closeMenus()">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">save</span>
                    <span class="dd-text">Save Now</span>
                    <span class="dd-hint" style="margin-left:auto; color:#5f6368;">Ctrl+S</span>
                  </div>

                  <!-- Download As -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subDownload)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">download</span><span class="dd-text" style="color:#1a73e8;">Download As</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subDownload class="sub-dropdown">
                       <div class="dd-item" (click)="exportFile('docx')"><span class="dd-text" style="color:#1a73e8;">MS Word (.DOCX)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.docx</span></div>
                       <div class="dd-item" (click)="openPasswordModal('docx')"><span class="dd-text">Password Protected MS Word (.docx)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.docx</span></div>
                       <div class="dd-item" (click)="exportFile('pdf')"><span class="dd-text" style="color:#1a73e8;">PDF (.PDF)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.pdf</span></div>
                       <div class="dd-item" (click)="openPasswordModal('pdf')"><span class="dd-text">Password-Protected PDF (.PDF)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.pdf</span></div>
                       <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subOtherFiles)">
                          <span class="dd-text">Other File Formats</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                          <div #subOtherFiles class="sub-dropdown">
                             <div class="dd-item" (click)="exportFile('html')"><span class="dd-text" style="color:#1a73e8;">Web Page (.html)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.html</span></div>
                             <div class="dd-item" (click)="exportFile('txt')"><span class="dd-text" style="color:#1a73e8;">Plain Text (.txt)</span><span class="dd-hint" style="margin-left:auto; color:#5f6368;">.txt</span></div>
                          </div>
                       </div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <!-- Document Versions -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subVersions)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">history</span><span class="dd-text" style="color:#1a73e8;">Document Versions</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subVersions class="sub-dropdown">
                       <div class="dd-item" (click)="showToast('Create Version')"><span class="material-symbols-outlined dd-icon" style="font-size:16px; color:#1a73e8;">history_toggle_off</span><span class="dd-text" style="color:#1a73e8;">Create Version</span></div>
                       <div class="dd-item" (click)="showVersionHistory()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">history</span><span class="dd-text">Version History...</span><span class="dd-hint">Ctrl+Alt+Shift+G</span></div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <!-- Mark As Final -->
                  <div class="dd-item" (click)="markAsFinalModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon">task_alt</span><span class="dd-text">Mark As Final</span></div>

                  <!-- Share -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subShare)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">share</span><span class="dd-text" style="color:#1a73e8;">Share</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subShare class="sub-dropdown">
                       <div class="dd-item" (click)="shareModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon" style="font-size:16px; color:#1a73e8;">group_add</span><span class="dd-text" style="color:#1a73e8;">Invite Collaborators</span></div>
                       <div class="dd-item" (click)="showToast('Share To Support')"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">support_agent</span><span class="dd-text">Share To Support</span></div>
                       <div class="dd-item" (click)="emailDoc()"><span class="material-symbols-outlined dd-icon" style="font-size:16px;">mail</span><span class="dd-text">Email As Attachment...</span></div>
                    </div>
                  </div>

                  <!-- Publish -->
                  <div class="dd-item" (click)="publishModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon">language</span><span class="dd-text">Publish</span></div>

                  <!-- Send for sign -->
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSign)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">draw</span><span class="dd-text" style="color:#1a73e8;">Send for sign</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subSign class="sub-dropdown">
                       <div class="dd-item" (click)="showToast('VMail Sign')" style="align-items:flex-start;">
                         <span class="material-symbols-outlined dd-icon" style="margin-top:2px; font-size:16px; color:#1a73e8;">draw</span>
                         <div style="display:flex; flex-direction:column;">
                           <span class="dd-text" style="color:#1a73e8;">Send For Sign Via VMail Sign</span>
                           <span style="font-size:10px; color:#5f6368; line-height:1; margin-top:2px;">Collect digital signatures using VMail Sign</span>
                         </div>
                       </div>
                       <div class="dd-item" (click)="showToast('Upload to Sign Services')" style="align-items:flex-start;">
                         <span class="material-symbols-outlined dd-icon" style="margin-top:2px; font-size:16px;">verified</span>
                         <div style="display:flex; flex-direction:column; flex:1;">
                           <span class="dd-text">Upload to Sign Services</span>
                           <span style="font-size:10px; color:#5f6368; line-height:1; margin-top:2px;">Collect signatures using your preferred<br>signature service</span>
                         </div>
                         <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px;">chevron_right</span>
                       </div>
                    </div>
                  </div>

                  <!-- Fill As Form -->
                  <div class="dd-item" (click)="fillFormModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon">list_alt</span><span class="dd-text">Fill As Form</span></div>

                  <div class="dd-sep"></div>

                  <!-- Page Setup -->
                  <div class="dd-item" (click)="pageSetupModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon">settings</span><span class="dd-text">Page Setup</span></div>

                  <!-- Print -->
                  <div class="dd-item" (click)="printDoc()"><span class="material-symbols-outlined dd-icon">print</span><span class="dd-text">Print</span><span class="dd-hint">Ctrl+P</span></div>

                  <div class="dd-item" (click)="showDetails()"><span class="material-symbols-outlined dd-icon">info</span><span class="dd-text">Document Properties</span></div>

                  <div class="dd-sep"></div>

                  <div class="dd-item" (click)="trashDoc()"><span class="material-symbols-outlined dd-icon">delete</span><span class="dd-text">Move To Trash</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('edit', $event)" [class.active]="activeMenu === 'edit'">
                Edit
                <div class="dropdown" *ngIf="activeMenu === 'edit'">
                  <div class="dd-item" (click)="exec('undo')"><span class="material-symbols-outlined dd-icon">undo</span><span class="dd-text">Undo</span><span class="dd-hint">Ctrl+Z</span></div>
                  <div class="dd-item" (click)="exec('redo')"><span class="material-symbols-outlined dd-icon">redo</span><span class="dd-text">Redo</span><span class="dd-hint">Ctrl+Y</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="showToast('Use Ctrl+X to cut')"><span class="material-symbols-outlined dd-icon">content_cut</span><span class="dd-text">Cut</span><span class="dd-hint">Ctrl+X</span></div>
                  <div class="dd-item" (click)="showToast('Use Ctrl+C to copy')"><span class="material-symbols-outlined dd-icon">content_copy</span><span class="dd-text">Copy</span><span class="dd-hint">Ctrl+C</span></div>
                  <div class="dd-item" (click)="showToast('Use Ctrl+V to paste')"><span class="material-symbols-outlined dd-icon">content_paste</span><span class="dd-text">Paste</span><span class="dd-hint">Ctrl+V</span></div>
                  <div class="dd-item" (click)="exec('selectAll')"><span class="dd-text">Select all</span><span class="dd-hint">Ctrl+A</span></div>
                  <div class="dd-item" (click)="exec('delete')"><span class="dd-text">Delete</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="exec('selectAll')"><span class="material-symbols-outlined dd-icon">select_all</span><span class="dd-text">Select all</span><span class="dd-hint">Ctrl+A</span></div>
                  <div class="dd-item" (click)="exec('delete')"><span class="material-symbols-outlined dd-icon">backspace</span><span class="dd-text">Delete</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="openFind()"><span class="material-symbols-outlined dd-icon">search</span><span class="dd-text">Find and replace</span><span class="dd-hint">Ctrl+F</span></div>
                  <div class="dd-item" (click)="exec('delete')"><span class="material-symbols-outlined dd-icon">backspace</span><span class="dd-text">Delete</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="openFind()"><span class="material-symbols-outlined dd-icon">search</span><span class="dd-text">Find and replace</span><span class="dd-hint">Ctrl+F</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('view', $event)" [class.active]="activeMenu === 'view'">
                View
                <div class="dropdown" *ngIf="activeMenu === 'view'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subDocView)">
                    <span class="material-symbols-outlined dd-icon" style="color: #1a73e8;">plagiarism</span>
                    <span class="dd-text" style="color: #1a73e8;">Document View : {{ documentViewType }}</span>
                    <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color: #1a73e8;">chevron_right</span>
                    <div #subDocView class="sub-dropdown">
                      <div class="dd-item" (click)="setDocumentView('Page View')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ documentViewType === 'Page View' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon" style="font-size:16px;">article</span><span class="dd-text" [style.color]="documentViewType === 'Page View' ? '#1a73e8' : ''">Page View</span></div>
                      <div class="dd-item" (click)="setDocumentView('Web View')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ documentViewType === 'Web View' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon" style="font-size:16px;">web</span><span class="dd-text" [style.color]="documentViewType === 'Web View' ? '#1a73e8' : ''">Web View at Width: 100%</span></div>
                    </div>
                  </div>
                  <div class="dd-item" (click)="toggleReaderView()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ isReaderView ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">menu_book</span><span class="dd-text">Reader View</span></div>
                  <div class="dd-item" (click)="toggleFullScreen()"><span class="material-symbols-outlined dd-icon">fullscreen</span><span class="dd-text">Full Screen</span><span class="dd-hint">Ctrl+F11</span></div>
                  
                  <div class="dd-sep"></div>
                  
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subZoom)">
                    <span class="material-symbols-outlined dd-icon" style="color: #1a73e8;">zoom_in</span>
                    <span class="dd-text" style="color: #1a73e8;">Zoom : {{ zoomLevel }}%</span>
                    <span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color: #1a73e8;">chevron_right</span>
                    <div #subZoom class="sub-dropdown">
                      <div class="dd-item" (click)="setZoom(200)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 200 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 200 ? '#1a73e8' : ''">200%</span></div>
                      <div class="dd-item" (click)="setZoom(175)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 175 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 175 ? '#1a73e8' : ''">175%</span></div>
                      <div class="dd-item" (click)="setZoom(150)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 150 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 150 ? '#1a73e8' : ''">150%</span></div>
                      <div class="dd-item" (click)="setZoom(125)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 125 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 125 ? '#1a73e8' : ''">125%</span></div>
                      <div class="dd-item" (click)="setZoom(100)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 100 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 100 ? '#1a73e8' : ''">100%</span></div>
                      <div class="dd-item" (click)="setZoom(75)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 75 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 75 ? '#1a73e8' : ''">75%</span></div>
                      <div class="dd-item" (click)="setZoom(50)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 50 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 50 ? '#1a73e8' : ''">50%</span></div>
                      <div class="dd-item" (click)="setZoom(25)"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ zoomLevel === 25 ? 'check' : '' }}</span><span class="dd-text" [style.color]="zoomLevel === 25 ? '#1a73e8' : ''">25%</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="fitWidth()"><span class="dd-text" style="padding-left: 32px;">Fit Width</span></div>
                      <div class="dd-item" (click)="fitPageToWindow()"><span class="dd-text" style="padding-left: 32px;">Fit Page To Window</span></div>
                    </div>
                  </div>
                  
                  <div class="dd-sep"></div>
                  
                  <div class="dd-item" (click)="toggleNavigator()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ showNavigator ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">format_list_bulleted</span><span class="dd-text">Navigator</span></div>
                  <div class="dd-item" (click)="toggleHideImages()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ hideImages ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">hide_image</span><span class="dd-text">Hide Images</span></div>
                  <div class="dd-item" (click)="toggleRuler()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ showRuler ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">straighten</span><span class="dd-text">Ruler</span></div>
                  
                  <div class="dd-item" (click)="toggleBookmarks()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ showBookmarks ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">bookmark</span><span class="dd-text">Bookmarks</span></div>
                  <div class="dd-item" (click)="toggleSmartGridLines()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ showSmartGridLines ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">grid_on</span><span class="dd-text">Smart Grid Lines</span></div>
                  <div class="dd-item" (click)="toggleObjectIndicator()"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ showObjectIndicator ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">anchor</span><span class="dd-text">Object Indicator</span></div>
                  <div class="dd-item" (click)="toggleFormatting()"><span class="material-symbols-outlined dd-icon">format_paragraph</span><span class="dd-text">Toggle Formatting ...</span><span class="dd-hint">Ctrl+Shift+8</span></div>
                  
                  <div class="dd-sep"></div>
                  
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subApp)">
                    <span class="material-symbols-outlined dd-icon" style="color: #1a73e8;">light_mode</span><span class="dd-text" style="color: #1a73e8;">Appearance</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color: #1a73e8;">chevron_right</span>
                    <div #subApp class="sub-dropdown">
                      <div class="dd-item" (click)="setAppearance('System Default')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ appearanceMode === 'System Default' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon" style="font-size:16px; color:#1a73e8;">computer</span><span class="dd-text" [style.color]="appearanceMode === 'System Default' ? '#1a73e8' : ''">System Default</span></div>
                      <div class="dd-item" (click)="setAppearance('Light Mode')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ appearanceMode === 'Light Mode' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon" style="font-size:16px;">light_mode</span><span class="dd-text" [style.color]="appearanceMode === 'Light Mode' ? '#1a73e8' : ''">Light Mode</span></div>
                      <div class="dd-item" (click)="setAppearance('Dark Mode')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ appearanceMode === 'Dark Mode' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon" style="font-size:16px;">dark_mode</span><span class="dd-text" [style.color]="appearanceMode === 'Dark Mode' ? '#1a73e8' : ''">Dark Mode <span style="color:#5f6368; font-size: 12px;">(Excluding editor area)</span></span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('More Settings')"><span class="dd-text" style="padding-left: 32px;">More Settings</span></div>
                    </div>
                  </div>
                  
                  <div class="dd-sep"></div>
                  
                  <div class="dd-item" (click)="showToast('More View Options')"><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More View Options ...</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('insert', $event)" [class.active]="activeMenu === 'insert'">
                Insert
                <div class="dropdown" *ngIf="activeMenu === 'insert'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subImageInsert)">
                    <span class="material-symbols-outlined dd-icon">image</span><span class="dd-text">Image</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subImageInsert class="sub-dropdown">
                      <div class="dd-item" (click)="imageInput.click(); closeMenus()"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">upload</span><span class="dd-text" style="color:#1a73e8;">Upload...</span></div>
                      <div class="dd-item" (click)="openImageModal('url')"><span class="material-symbols-outlined dd-icon">link</span><span class="dd-text">Insert a URL</span></div>
                      <div class="dd-item" (click)="openImageModal('workdrive')"><span class="material-symbols-outlined dd-icon">cloud</span><span class="dd-text">Pick From WorkDrive</span></div>
                      <div class="dd-item" (click)="openImageModal('library')"><span class="material-symbols-outlined dd-icon">photo_library</span><span class="dd-text">My Library</span></div>
                      <div class="dd-item" (click)="openImageModal('gphotos')"><span class="material-symbols-outlined dd-icon">photo_camera</span><span class="dd-text">Google Photos</span></div>
                      <div class="dd-item" (click)="openImageModal('flickr')"><span class="material-symbols-outlined dd-icon">camera</span><span class="dd-text">Flickr</span></div>
                      <div class="dd-item" (click)="openImageModal('web')"><span class="material-symbols-outlined dd-icon">public</span><span class="dd-text">Pick From the Web</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subTableInsert)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">grid_on</span><span class="dd-text" style="color:#1a73e8;">Table</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subTableInsert class="sub-dropdown" style="min-width: 250px;">
                      <div style="padding: 12px;" (mouseleave)="hoveredTableRow = 0; hoveredTableCol = 0">
                        <div style="margin-bottom: 8px; font-size: 13px; font-weight: 600;">Table</div>
                        <div style="display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px;">
                          <ng-container *ngFor="let r of tableGridRows">
                            <div *ngFor="let c of tableGridCols" 
                                 [style.border]="(r <= hoveredTableRow && c <= hoveredTableCol) ? '1px solid #1a73e8' : '1px solid #dadce0'" 
                                 [style.background]="(r <= hoveredTableRow && c <= hoveredTableCol) ? '#e8f0fe' : 'transparent'" 
                                 style="height: 16px; cursor: pointer;" 
                                 (mouseenter)="hoveredTableRow = r; hoveredTableCol = c"
                                 (click)="insertGridTable(hoveredTableRow + 1, hoveredTableCol + 1)">
                            </div>
                          </ng-container>
                        </div>
                        <div style="text-align: center; font-size: 11px; margin-top: 8px; color: #5f6368; font-weight: 600;">{{ hoveredTableRow + 1 }} x {{ hoveredTableCol + 1 }}</div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="insertTable()"><span class="material-symbols-outlined dd-icon">table_view</span><span class="dd-text">Specify Rows And Columns</span></div>
                      <div class="dd-item" style="color: #ccc; cursor: default;"><span class="material-symbols-outlined dd-icon" style="color: #ccc;">text_snippet</span><span class="dd-text">Convert Text To Table</span></div>
                      <div class="dd-item" (click)="showToast('Insert A New Spreadsheet')"><span class="material-symbols-outlined dd-icon">grid_on</span><span class="dd-text">Insert A New Spreadsheet</span></div>
                      <div class="dd-item" (click)="showToast('Pick An Existing Spreadsheet')"><span class="material-symbols-outlined dd-icon">library_add</span><span class="dd-text">Pick An Existing Spreadsheet</span></div>
                    </div>
                  </div>

                  <div class="dd-item" (click)="insertTextBox()"><span class="material-symbols-outlined dd-icon">text_fields</span><span class="dd-text">Text Box</span></div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subQuickParts)">
                    <span class="material-symbols-outlined dd-icon">post_add</span><span class="dd-text">Quick Parts</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subQuickParts class="sub-dropdown">
                      <div class="dd-item" (click)="insertQuickPart()"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">post_add</span><span class="dd-text" style="color:#1a73e8;">Insert Signature Block</span></div>
                      <div style="font-size: 11px; font-weight: bold; color: #5f6368; padding: 4px 16px; margin-top: 8px; border-bottom: 1px solid #eee; padding-bottom: 8px;">Use selection to</div>
                      <div class="dd-item" (click)="saveQuickPart()"><span class="material-symbols-outlined dd-icon">save</span><span class="dd-text">Save to Quick Parts Gallery</span></div>
                      <div class="dd-item" style="color: #ccc; cursor: default;"><span class="material-symbols-outlined dd-icon" style="color: #ccc;">add_circle</span><span class="dd-text">Create New</span></div>
                      <div class="dd-item" (click)="showToast('Manage Quick Parts')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">settings</span><span class="dd-text" style="color:#1a73e8;">Manage Quick Parts</span></div>
                    </div>
                  </div>

                  <div class="dd-item" (click)="insertDrawing()"><span class="material-symbols-outlined dd-icon">draw</span><span class="dd-text">Drawings...</span></div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSymbols)">
                    <span class="material-symbols-outlined dd-icon">special_character</span><span class="dd-text">Symbols...</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subSymbols class="sub-dropdown" style="width: 200px; padding: 12px; cursor: default;">
                      <div style="font-size: 11px; font-weight: bold; color: #5f6368; margin-bottom: 8px;">Common Symbols</div>
                      <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 4px; text-align: center;">
                        <span class="symbol-btn" (click)="insertText('©')">©</span>
                        <span class="symbol-btn" (click)="insertText('®')">®</span>
                        <span class="symbol-btn" (click)="insertText('™')">™</span>
                        <span class="symbol-btn" (click)="insertText('€')">€</span>
                        <span class="symbol-btn" (click)="insertText('£')">£</span>
                        <span class="symbol-btn" (click)="insertText('¥')">¥</span>
                        <span class="symbol-btn" (click)="insertText('•')">•</span>
                        <span class="symbol-btn" (click)="insertText('✓')">✓</span>
                        <span class="symbol-btn" (click)="insertText('×')">×</span>
                        <span class="symbol-btn" (click)="insertText('÷')">÷</span>
                        <span class="symbol-btn" (click)="insertText('±')">±</span>
                        <span class="symbol-btn" (click)="insertText('∞')">∞</span>
                        <span class="symbol-btn" (click)="insertText('∑')">∑</span>
                        <span class="symbol-btn" (click)="insertText('∆')">∆</span>
                        <span class="symbol-btn" (click)="insertText('π')">π</span>
                      </div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subEquations)">
                    <span class="material-symbols-outlined dd-icon">calculate</span><span class="dd-text">Equation...</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subEquations class="sub-dropdown" style="width: 220px; padding: 8px;">
                      <div class="dd-item" (click)="insertText('A = πr²')"><span class="dd-text">Area of Circle (A = πr²)</span></div>
                      <div class="dd-item" (click)="insertText('a² + b² = c²')"><span class="dd-text">Pythagorean (a² + b² = c²)</span></div>
                      <div class="dd-item" (click)="insertText('x = (-b ± √(b² - 4ac)) / 2a')"><span class="dd-text">Quadratic Formula</span></div>
                      <div class="dd-item" (click)="insertText('E = mc²')"><span class="dd-text">Mass-Energy Equivalence</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subChart)">
                    <span class="material-symbols-outlined dd-icon">bar_chart</span><span class="dd-text">Chart</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subChart class="sub-dropdown" style="min-width: 220px;">
                      <div style="font-size: 11px; font-weight: bold; color: #5f6368; padding: 4px 16px; margin-top: 4px; border-bottom: 1px solid #eee;">Bar</div>
                      <div style="padding: 8px 16px; display: flex; gap: 16px;">
                        <span class="material-symbols-outlined" style="color:#1a73e8; font-size:24px; cursor:pointer;" (click)="openChartModal('Bar')">bar_chart</span>
                        <span class="material-symbols-outlined" style="color:#34a853; font-size:24px; cursor:pointer;" (click)="openChartModal('Stacked Bar')">stacked_bar_chart</span>
                        <span class="material-symbols-outlined" style="color:#1a73e8; font-size:24px; cursor:pointer;" (click)="openChartModal('Waterfall')">waterfall_chart</span>
                      </div>
                      <div style="font-size: 11px; font-weight: bold; color: #5f6368; padding: 4px 16px; margin-top: 4px; border-bottom: 1px solid #eee;">Line</div>
                      <div style="padding: 8px 16px; display: flex; gap: 16px;">
                        <span class="material-symbols-outlined" style="color:#34a853; font-size:24px; cursor:pointer;" (click)="openChartModal('Line')">show_chart</span>
                        <span class="material-symbols-outlined" style="color:#1a73e8; font-size:24px; cursor:pointer;" (click)="openChartModal('Stacked Line')">stacked_line_chart</span>
                      </div>
                      <div style="font-size: 11px; font-weight: bold; color: #5f6368; padding: 4px 16px; margin-top: 4px; border-bottom: 1px solid #eee;">Pie</div>
                      <div style="padding: 8px 16px; display: flex; gap: 16px;">
                        <span class="material-symbols-outlined" style="color:#fbbc04; font-size:24px; cursor:pointer;" (click)="openChartModal('Pie')">pie_chart</span>
                        <span class="material-symbols-outlined" style="color:#1a73e8; font-size:24px; cursor:pointer;" (click)="openChartModal('Donut')">donut_large</span>
                      </div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMedia)">
                    <span class="material-symbols-outlined dd-icon">play_circle</span><span class="dd-text">Media</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subMedia class="sub-dropdown">
                      <div class="dd-item" (click)="videoInput.click()"><span class="material-symbols-outlined dd-icon">upload</span><span class="dd-text">Upload Video</span></div>
                      <div class="dd-item" (click)="insertVideo()"><span class="material-symbols-outlined dd-icon">movie</span><span class="dd-text">Video from URL</span></div>
                      <div class="dd-item" (click)="audioInput.click()"><span class="material-symbols-outlined dd-icon">upload</span><span class="dd-text">Upload Audio</span></div>
                      <div class="dd-item" (click)="insertAudio()"><span class="material-symbols-outlined dd-icon">audio_file</span><span class="dd-text">Audio from URL</span></div>
                    </div>
                  </div>
                  
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subTOC)">
                    <span class="material-symbols-outlined dd-icon">toc</span><span class="dd-text">Table of Contents...</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subTOC class="sub-dropdown">
                      <div class="dd-item" (click)="insertTOC('plain')"><span class="material-symbols-outlined dd-icon">list_alt</span><span class="dd-text">Plain Text</span></div>
                      <div class="dd-item" (click)="insertTOC('links')"><span class="material-symbols-outlined dd-icon">link</span><span class="dd-text">With Page Links</span></div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertLink()"><span class="material-symbols-outlined dd-icon">link</span><span class="dd-text">Link...</span><span class="dd-hint">Ctrl+K</span></div>
                  <div class="dd-item" (click)="showToast('Bookmark')"><span class="material-symbols-outlined dd-icon">bookmark</span><span class="dd-text">Bookmark...</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertComment()"><span class="material-symbols-outlined dd-icon">comment</span><span class="dd-text">Comment</span><span class="dd-hint">Ctrl+Alt+M</span></div>

                  <div class="dd-sep"></div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subPageBreak)">
                    <span class="material-symbols-outlined dd-icon">insert_page_break</span><span class="dd-text">Page Break</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subPageBreak class="sub-dropdown">
                      <div class="dd-item" (click)="insertBreak()"><span class="material-symbols-outlined dd-icon">insert_page_break</span><span class="dd-text">Page Break</span></div>
                      <div class="dd-item" (click)="showToast('Section Break')"><span class="material-symbols-outlined dd-icon">vertical_split</span><span class="dd-text">Section Break</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subPageElements)">
                    <span class="material-symbols-outlined dd-icon">contact_page</span><span class="dd-text">Page Elements</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subPageElements class="sub-dropdown">
                      <div class="dd-item" (click)="insertHeader()"><span class="material-symbols-outlined dd-icon">horizontal_rule</span><span class="dd-text">Header</span></div>
                      <div class="dd-item" (click)="insertFooter()"><span class="material-symbols-outlined dd-icon">horizontal_rule</span><span class="dd-text">Footer</span></div>
                      <div class="dd-item" (click)="insertPageNumbers()"><span class="material-symbols-outlined dd-icon">numbers</span><span class="dd-text">Page Numbers</span></div>
                    </div>
                  </div>

                  <div class="dd-item" (click)="exec('insertHorizontalRule')"><span class="material-symbols-outlined dd-icon">horizontal_rule</span><span class="dd-text">Horizontal Line</span></div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subQrBarcode)">
                    <span class="material-symbols-outlined dd-icon">qr_code_2</span><span class="dd-text">QR & Barcode</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subQrBarcode class="sub-dropdown">
                      <div class="dd-item" (click)="insertQRCode()"><span class="material-symbols-outlined dd-icon">qr_code</span><span class="dd-text">QR Code</span></div>
                      <div class="dd-item" (click)="insertBarcode()"><span class="material-symbols-outlined dd-icon">view_headline</span><span class="dd-text">Barcode</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subDropCap)">
                    <span class="material-symbols-outlined dd-icon">format_shapes</span><span class="dd-text">Drop Cap</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subDropCap class="sub-dropdown">
                      <div class="dd-item" (click)="applyDropCap('None')"><span class="material-symbols-outlined dd-icon">close</span><span class="dd-text">None</span></div>
                      <div class="dd-item" (click)="applyDropCap('Dropped')"><span class="material-symbols-outlined dd-icon">format_shapes</span><span class="dd-text">Dropped</span></div>
                      <div class="dd-item" (click)="applyDropCap('In margin')"><span class="material-symbols-outlined dd-icon">format_shapes</span><span class="dd-text">In margin</span></div>
                    </div>
                  </div>

                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertCode()"><span class="material-symbols-outlined dd-icon">code</span><span class="dd-text">Code</span></div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSignature)">
                    <span class="material-symbols-outlined dd-icon">draw</span><span class="dd-text">Signature</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subSignature class="sub-dropdown">
                      <div class="dd-item" (click)="insertDrawing()"><span class="material-symbols-outlined dd-icon">draw</span><span class="dd-text">Draw Signature</span></div>
                      <div class="dd-item" (click)="signatureInput.click()"><span class="material-symbols-outlined dd-icon">upload</span><span class="dd-text">Upload Signature</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subReferences)">
                    <span class="material-symbols-outlined dd-icon">library_books</span><span class="dd-text">References</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subReferences class="sub-dropdown">
                      <div class="dd-item" (click)="showToast('Footnote')"><span class="material-symbols-outlined dd-icon">format_list_numbered</span><span class="dd-text">Footnote</span></div>
                      <div class="dd-item" (click)="showToast('Endnote')"><span class="material-symbols-outlined dd-icon">format_list_numbered_rtl</span><span class="dd-text">Endnote</span></div>
                    </div>
                  </div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('format', $event)" [class.active]="activeMenu === 'format'">
                Format
                <div class="dropdown" *ngIf="activeMenu === 'format'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subTextFormat)">
                    <span class="material-symbols-outlined dd-icon">text_format</span><span class="dd-text">Text Formatting</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subTextFormat class="sub-dropdown" style="min-width: 250px;">
                      <div class="dd-item" (click)="exec('bold')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_bold</span><span class="dd-text" style="color:#1a73e8;">Bold</span><span class="dd-hint">Ctrl+B</span></div>
                      <div class="dd-item" (click)="exec('italic')"><span class="material-symbols-outlined dd-icon">format_italic</span><span class="dd-text">Italic</span><span class="dd-hint">Ctrl+I</span></div>
                      <div class="dd-item" (click)="exec('underline')"><span class="material-symbols-outlined dd-icon">format_underlined</span><span class="dd-text">Underline</span><span class="dd-hint">Ctrl+U</span></div>
                      <div class="dd-item" (click)="exec('strikeThrough')"><span class="material-symbols-outlined dd-icon">strikethrough_s</span><span class="dd-text">Strikethrough</span><span class="dd-hint">Ctrl+Shift+X</span></div>
                      <div class="dd-item" (click)="exec('superscript')"><span class="material-symbols-outlined dd-icon">superscript</span><span class="dd-text">Superscript</span><span class="dd-hint">Ctrl+.</span></div>
                      <div class="dd-item" (click)="exec('subscript')"><span class="material-symbols-outlined dd-icon">subscript</span><span class="dd-text">Subscript</span><span class="dd-hint">Ctrl+,</span></div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subChangeCase)">
                         <span class="material-symbols-outlined dd-icon">match_case</span><span class="dd-text">Change Case</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subChangeCase class="sub-dropdown">
                            <div class="dd-item" (click)="showToast('lowercase')"><span class="dd-text">lowercase</span></div>
                            <div class="dd-item" (click)="showToast('UPPERCASE')"><span class="dd-text">UPPERCASE</span></div>
                            <div class="dd-item" (click)="showToast('Title Case')"><span class="dd-text">Title Case</span></div>
                         </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subCharSpace)">
                         <span class="material-symbols-outlined dd-icon">format_letter_spacing</span><span class="dd-text">Character Spacing</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subCharSpace class="sub-dropdown">
                            <div class="dd-item" (click)="showToast('Normal')"><span class="dd-text">Normal</span></div>
                            <div class="dd-item" (click)="showToast('Expanded')"><span class="dd-text">Expanded</span></div>
                            <div class="dd-item" (click)="showToast('Condensed')"><span class="dd-text">Condensed</span></div>
                         </div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subFontFamily)">
                         <span class="material-symbols-outlined dd-icon">font_download</span><span class="dd-text">Font Family: <span style="color:#1a73e8;">Roboto</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subFontFamily class="sub-dropdown" style="max-height: 250px; overflow-y: auto;">
                            <div class="dd-item" *ngFor="let f of fonts" (click)="execVal('fontName', f)"><span class="dd-text" [style.font-family]="f">{{f}}</span></div>
                         </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subFontSize)">
                         <span class="material-symbols-outlined dd-icon">format_size</span><span class="dd-text">Font Size: <span style="color:#1a73e8;">12 pt</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subFontSize class="sub-dropdown" style="max-height: 250px; overflow-y: auto;">
                            <div class="dd-item" *ngFor="let size of [8,9,10,11,12,14,18,24,30,36]" (click)="execVal('fontSize', size.toString())"><span class="dd-text">{{size}} pt</span></div>
                         </div>
                      </div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subParaStyle)">
                    <span class="material-symbols-outlined dd-icon">format_shapes</span><span class="dd-text">Paragraph Style: <span style="color:#1a73e8;">Normal</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subParaStyle class="sub-dropdown" style="min-width: 250px;">
                      <div class="dd-item" (click)="execVal('formatBlock', 'P')"><span class="dd-text" style="color:#1a73e8;">Normal</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="showToast('Title applied')"><span class="dd-text" style="color:#1a73e8; font-weight: bold; font-size: 16px;">Title</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="showToast('Subtitle applied')"><span class="dd-text" style="font-style: italic;">Subtitle</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H1')"><span class="dd-text" style="font-weight: bold; font-size: 16px;">Heading 1</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H2')"><span class="dd-text" style="font-weight: bold; font-size: 15px;">Heading 2</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H3')"><span class="dd-text" style="font-weight: bold; font-size: 14px;">Heading 3</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H4')"><span class="dd-text" style="font-weight: bold; font-style: italic;">Heading 4</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H5')"><span class="dd-text" style="background:#444; color:#fff; padding:2px 4px;">Heading 5</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'H6')"><span class="dd-text" style="text-decoration: underline; color:#1a73e8;">Heading 6</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-item" (click)="execVal('formatBlock', 'BLOCKQUOTE')"><span class="dd-text" style="background:#e8f0fe; padding:2px 4px;">Quote</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Set Next Paragraph Style')"><span class="material-symbols-outlined dd-icon">text_format</span><span class="dd-text">Set Next Paragraph Style</span></div>
                      <div class="dd-item" (click)="showToast('Add From Style Library')"><span class="material-symbols-outlined dd-icon">library_add</span><span class="dd-text">Add From Style Library</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subAlign)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_align_left</span><span class="dd-text" style="color:#1a73e8;">Align</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subAlign class="sub-dropdown">
                      <div class="dd-item" (click)="exec('justifyLeft')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_align_left</span><span class="dd-text" style="color:#1a73e8;">Align Left</span><span class="dd-hint">Ctrl+Shift+L</span></div>
                      <div class="dd-item" (click)="exec('justifyCenter')"><span class="material-symbols-outlined dd-icon">format_align_center</span><span class="dd-text">Align Center</span><span class="dd-hint">Ctrl+Shift+E</span></div>
                      <div class="dd-item" (click)="exec('justifyRight')"><span class="material-symbols-outlined dd-icon">format_align_right</span><span class="dd-text">Align Right</span><span class="dd-hint">Ctrl+Shift+R</span></div>
                      <div class="dd-item" (click)="exec('justifyFull')"><span class="material-symbols-outlined dd-icon">format_align_justify</span><span class="dd-text">Justify</span><span class="dd-hint">Ctrl+Shift+J</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subLineSpacing)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_line_spacing</span><span class="dd-text" style="color:#1a73e8;">Line & Paragraph Spacing</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subLineSpacing class="sub-dropdown" style="min-width: 280px;">
                      <div class="dd-item" (click)="showToast('Line Spacing 1.0')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">format_line_spacing</span><span class="dd-text" style="color:#1a73e8;">1.0 (Single)</span></div>
                      <div class="dd-item" (click)="showToast('Line Spacing 1.2')"><span class="material-symbols-outlined dd-icon">check</span><span class="material-symbols-outlined dd-icon">format_line_spacing</span><span class="dd-text">1.2 (Normal)</span></div>
                      <div class="dd-item" (click)="showToast('Line Spacing 1.5')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">format_line_spacing</span><span class="dd-text">1.5</span></div>
                      <div class="dd-item" (click)="showToast('Line Spacing 2.0')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">format_line_spacing</span><span class="dd-text">2.0 (Double)</span></div>
                      <div class="dd-item" (click)="showToast('Custom Spacing')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">linear_scale</span><span class="dd-text">Custom Spacing</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subCharSpace2)">
                         <span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">format_letter_spacing</span><span class="dd-text">Character Spacing</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subCharSpace2 class="sub-dropdown">
                            <div class="dd-item" (click)="showToast('Normal')"><span class="dd-text">Normal</span></div>
                            <div class="dd-item" (click)="showToast('Expanded')"><span class="dd-text">Expanded</span></div>
                            <div class="dd-item" (click)="showToast('Condensed')"><span class="dd-text">Condensed</span></div>
                         </div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Add Space Before')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">vertical_align_bottom</span><span class="dd-text">Add Space Before Paragraph</span></div>
                      <div class="dd-item" (click)="showToast('Remove Space After')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">vertical_align_top</span><span class="dd-text">Remove Space After Paragraph</span></div>
                      <div class="dd-item" (click)="showToast('More Spacing Options')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More Paragraph Spacing Options</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Widow Control')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">subject</span><span class="dd-text">Enable Widow/Orphan Control</span></div>
                      <div class="dd-item" (click)="showToast('Keep Lines')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">dehaze</span><span class="dd-text">Keep Lines Together</span></div>
                      <div class="dd-item" (click)="showToast('Page Break Before')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">insert_page_break</span><span class="dd-text">Page Break Before</span></div>
                      <div class="dd-item" (click)="showToast('Keep With Next')"><span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span><span class="material-symbols-outlined dd-icon">low_priority</span><span class="dd-text">Keep With Next</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subList)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_list_bulleted</span><span class="dd-text" style="color:#1a73e8;">List</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subList class="sub-dropdown">
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subBullets)">
                         <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">format_list_bulleted</span><span class="dd-text" style="color:#1a73e8;">Bulleted List</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                         <div #subBullets class="sub-dropdown" style="min-width: 220px;">
                            <div style="padding: 12px;">
                               <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 4px;">
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span style="font-size:10px; font-weight:bold; color:#5f6368;">NONE</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">fiber_manual_record</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">panorama_fish_eye</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">stop</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">check_box_outline_blank</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">play_arrow</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">change_history</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">star</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">adjust</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px; color:#1a73e8;">send</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px;">check</span></div>
                                  <div style="border: 1px solid #dadce0; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; border-radius: 4px;" (click)="exec('insertUnorderedList')"><span class="material-symbols-outlined" style="font-size: 16px; color:#fbbc04;">widgets</span></div>
                               </div>
                            </div>
                            <div class="dd-sep"></div>
                            <div class="dd-item" (click)="showToast('Use Symbol')"><span class="dd-text">Use Symbol</span></div>
                            <div class="dd-item" (click)="showToast('Use Image')"><span class="dd-text">Use Image</span></div>
                         </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subNumbers)">
                         <span class="material-symbols-outlined dd-icon">format_list_numbered</span><span class="dd-text">Numbered List</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subNumbers class="sub-dropdown">
                            <div class="dd-item" (click)="exec('insertOrderedList')"><span class="dd-text">1, 2, 3</span></div>
                            <div class="dd-item" (click)="exec('insertOrderedList')"><span class="dd-text">a, b, c</span></div>
                            <div class="dd-item" (click)="exec('insertOrderedList')"><span class="dd-text">i, ii, iii</span></div>
                         </div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subChecklist)">
                         <span class="material-symbols-outlined dd-icon">checklist</span><span class="dd-text">Checklist</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                         <div #subChecklist class="sub-dropdown">
                            <div class="dd-item" (click)="showToast('Insert Checklist')"><span class="dd-text">Standard</span></div>
                            <div class="dd-item" (click)="showToast('Insert Strike-through Checklist')"><span class="dd-text">Strikethrough on check</span></div>
                         </div>
                      </div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subIndent)">
                    <span class="material-symbols-outlined dd-icon">format_indent_increase</span><span class="dd-text">Indent</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subIndent class="sub-dropdown">
                      <div class="dd-item" (click)="exec('indent')"><span class="material-symbols-outlined dd-icon">format_indent_increase</span><span class="dd-text">Increase Indent</span></div>
                      <div class="dd-item" (click)="exec('outdent')"><span class="material-symbols-outlined dd-icon">format_indent_decrease</span><span class="dd-text">Decrease Indent</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subTextDir)">
                    <span class="material-symbols-outlined dd-icon">format_textdirection_l_to_r</span><span class="dd-text">Text Direction</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subTextDir class="sub-dropdown">
                      <div class="dd-item" (click)="showToast('Left to Right')"><span class="material-symbols-outlined dd-icon">format_textdirection_l_to_r</span><span class="dd-text">Left-to-Right</span></div>
                      <div class="dd-item" (click)="showToast('Right to Left')"><span class="material-symbols-outlined dd-icon">format_textdirection_r_to_l</span><span class="dd-text">Right-to-Left</span></div>
                    </div>
                  </div>

                  <div class="dd-item" (click)="showToast('Borders Options')"><span class="material-symbols-outlined dd-icon">border_all</span><span class="dd-text">Borders and Shading Options</span></div>

                  <div class="dd-sep"></div>

                  <div class="dd-item" (click)="showToast('Format Painter')"><span class="material-symbols-outlined dd-icon">format_paint</span><span class="dd-text">Format Painter</span><span class="dd-hint">Ctrl+Shift+C</span></div>
                  <div class="dd-item" (click)="exec('removeFormat')"><span class="material-symbols-outlined dd-icon">format_clear</span><span class="dd-text">Clear Formatting</span><span class="dd-hint">Ctrl+\</span></div>

                  <div class="dd-sep"></div>

                  <div class="dd-item" (click)="showToast('More Format Options')"><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More Format Options...</span></div>
                  <div class="dd-item" (click)="showToast('More Paragraph Options')"><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More Paragraph Options...</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('design', $event)" [class.active]="activeMenu === 'design'">
                Design
                <div class="dropdown" *ngIf="activeMenu === 'design'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subCurrentDesign)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">palette</span><span class="dd-text">Current Design: <span style="color:#1a73e8;">{{ activeCurrentDesign }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subCurrentDesign class="sub-dropdown" style="min-width: 250px;">
                      <div style="padding: 16px; text-align: center;">
                        <div style="font-weight: 600; color: #5f6368; margin-bottom: 8px;">{{ activeCurrentDesign }}</div>
                        <div style="border: 1px solid #dadce0; padding: 16px; height: 180px; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: flex-start; overflow: hidden; pointer-events: none;">
                          <h1 style="color: #1a73e8; margin: 0 0 8px 0; font-size: 20px;">TITLE</h1>
                          <div style="color: #5f6368; font-size: 10px; font-style: italic; margin-bottom: 8px;">Subtitle</div>
                          <h2 style="margin: 0 0 4px 0; font-size: 16px;">Heading 1</h2>
                          <div style="font-size: 8px; color: #5f6368; text-align: left; line-height: 1.2;">Lorem ipsum dolor sit amet adipiscing elit, fusce odio laoreet eleifend.</div>
                        </div>
                        <div style="display: flex; justify-content: center; gap: 4px; margin-top: 8px;">
                          <div style="width: 6px; height: 6px; border-radius: 50%; background: #5f6368;"></div>
                          <div style="width: 6px; height: 6px; border-radius: 50%; background: #dadce0;"></div>
                        </div>
                        <div style="margin-top: 12px;">
                          <button style="width: 100%; padding: 6px; background: #fff; border: 1px solid #dadce0; color: #1a73e8; font-weight: 600; border-radius: 4px; cursor: pointer;">Save As...</button>
                        </div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" style="color:#ccc; cursor:default;"><span class="material-symbols-outlined dd-icon" style="color:#ccc;">edit_document</span><span class="dd-text">Rename Design</span></div>
                      <div class="dd-item" (click)="showToast('Set As Default Design')"><span class="material-symbols-outlined dd-icon">check_circle</span><span class="dd-text">Set As Default Design</span></div>
                      <div class="dd-item" (click)="showToast('Reset To Default')"><span class="material-symbols-outlined dd-icon">settings_backup_restore</span><span class="dd-text">Reset To Default</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subDesignGallery)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">style</span><span class="dd-text" style="color:#1a73e8;">Design Gallery</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subDesignGallery class="sub-dropdown">
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMyDesigns)">
                        <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">person</span><span class="dd-text" style="color:#1a73e8;">My Designs <span style="color:#8ab4f8; margin-left: 4px;">0</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                        <div #subMyDesigns class="sub-dropdown"><div class="dd-item" style="color:#9aa0a6;"><span class="dd-text">No Custom Designs Found</span></div></div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subOrgDesigns)">
                        <span class="material-symbols-outlined dd-icon" style="color:#ccc;">domain</span><span class="dd-text" style="color:#ccc;">Org Designs <span style="color:#dadce0; margin-left: 4px;">0</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#ccc;">chevron_right</span>
                        <div #subOrgDesigns class="sub-dropdown"><div class="dd-item" style="color:#9aa0a6;"><span class="dd-text">No Org Designs Found</span></div></div>
                      </div>
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subPresetDesigns)">
                        <span class="material-symbols-outlined dd-icon">check</span><span class="material-symbols-outlined dd-icon">tune</span><span class="dd-text">Preset Designs <span style="color:#1a73e8; margin-left: 4px;">19</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                        <div #subPresetDesigns class="sub-dropdown" style="max-height: 350px; overflow-y: auto;">
                          <div class="dd-item" (click)="applyDesign('The Writer')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'The Writer' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'The Writer' ? '#1a73e8' : '#000'">The Writer</span></div>
                          <div class="dd-item" (click)="applyDesign('Modern Report')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'Modern Report' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'Modern Report' ? '#1a73e8' : '#000'">Modern Report</span></div>
                          <div class="dd-item" (click)="applyDesign('Newsletter')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'Newsletter' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'Newsletter' ? '#1a73e8' : '#000'">Newsletter</span></div>
                          <div class="dd-item" (click)="applyDesign('Academic Paper')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'Academic Paper' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'Academic Paper' ? '#1a73e8' : '#000'">Academic Paper</span></div>
                          <div class="dd-item" (click)="applyDesign('Business Letter')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'Business Letter' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'Business Letter' ? '#1a73e8' : '#000'">Business Letter</span></div>
                          <div class="dd-item" (click)="applyDesign('Creative Story')"><span class="material-symbols-outlined dd-icon" [style.color]="activeCurrentDesign === 'Creative Story' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeCurrentDesign === 'Creative Story' ? '#1a73e8' : '#000'">Creative Story</span></div>
                          <div class="dd-sep"></div>
                          <div class="dd-item" (click)="showToast('Browse 13 more Presets...')"><span class="dd-text">Browse more...</span></div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subFontSet)">
                    <span class="material-symbols-outlined dd-icon">text_fields</span><span class="dd-text">Font Set</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subFontSet class="sub-dropdown" style="max-height: 350px; overflow-y: auto;">
                      <div class="dd-item" (click)="applyFontSet('Cabin')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Cabin' ? '#1a73e8' : '#000'">Cabin (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Istok Web')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Istok Web' ? '#1a73e8' : '#000'">Istok Web (Headings)</span><span class="dd-text" style="color:#5f6368;">Istok Web (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('PT Sans')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'PT Sans' ? '#1a73e8' : '#000'">PT Sans (Headings)</span><span class="dd-text" style="color:#5f6368;">Pontano Sans (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Work Sans')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Work Sans' ? '#1a73e8' : '#000'">Work Sans (Headings)</span><span class="dd-text" style="color:#5f6368;">Work Sans (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Roboto')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Roboto' ? '#1a73e8' : '#000'">Roboto (Headings)</span><span class="dd-text" style="color:#5f6368;">Roboto (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Rokkitt')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Rokkitt' ? '#1a73e8' : '#000'">Rokkitt (Headings)</span><span class="dd-text" style="color:#5f6368;">Rokkitt (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Quicksand')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Quicksand' ? '#1a73e8' : '#000'">Quicksand (Headings)</span><span class="dd-text" style="color:#5f6368;">Quicksand (Body)</span></div></div>
                      <div class="dd-item" (click)="applyFontSet('Source Sans Pro')"><div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeFontSet === 'Source Sans Pro' ? '#1a73e8' : '#000'">Source Sans Pro (Headings)</span><span class="dd-text" style="color:#5f6368;">Source Sans Pro (Body)</span></div></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Add Font Set')"><span class="material-symbols-outlined dd-icon" style="color:#34a853;">add_circle</span><span class="dd-text">Add Font Set</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subColorSet)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">color_lens</span><span class="dd-text">Color Set: <span style="color:#1a73e8;">{{ activeColorSet }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subColorSet class="sub-dropdown" style="max-height: 350px; overflow-y: auto;">
                      <div class="dd-item" (click)="applyColorSet('Brushstrokes')"><div style="display:flex; flex-direction:column; gap:4px;"><span class="dd-text" [style.color]="activeColorSet === 'Brushstrokes' ? '#1a73e8' : '#000'">Brushstrokes</span><div style="display:flex; gap:2px;"><div style="width:12px; height:12px; background:#5a4325;"></div><div style="width:12px; height:12px; background:#e0ecf8;"></div><div style="width:12px; height:12px; background:#a39171;"></div><div style="width:12px; height:12px; background:#f47c6b;"></div><div style="width:12px; height:12px; background:#475765;"></div><div style="width:12px; height:12px; background:#7b9ea3;"></div><div style="width:12px; height:12px; background:#b2d235;"></div></div></div></div>
                      <div class="dd-item" (click)="applyColorSet('bold')"><div style="display:flex; flex-direction:column; gap:4px;"><span class="dd-text" [style.color]="activeColorSet === 'bold' ? '#1a73e8' : '#000'">bold</span><div style="display:flex; gap:2px;"><div style="width:12px; height:12px; background:#333;"></div><div style="width:12px; height:12px; background:#f2e6d6;"></div><div style="width:12px; height:12px; background:#e85d50;"></div><div style="width:12px; height:12px; background:#00bcd4;"></div><div style="width:12px; height:12px; background:#ff9800;"></div><div style="width:12px; height:12px; background:#cddc39;"></div><div style="width:12px; height:12px; background:#e91e63;"></div></div></div></div>
                      <div class="dd-item" (click)="applyColorSet('Treasury')"><div style="display:flex; flex-direction:column; gap:4px;"><span class="dd-text" [style.color]="activeColorSet === 'Treasury' ? '#1a73e8' : '#000'">Treasury</span><div style="display:flex; gap:2px;"><div style="width:12px; height:12px; background:#3e1f1f;"></div><div style="width:12px; height:12px; background:#f5da55;"></div><div style="width:12px; height:12px; background:#9eabc1;"></div><div style="width:12px; height:12px; background:#b274a7;"></div><div style="width:12px; height:12px; background:#a78869;"></div><div style="width:12px; height:12px; background:#8b9d82;"></div><div style="width:12px; height:12px; background:#6b7c84;"></div></div></div></div>
                      <div class="dd-item" (click)="applyColorSet('Default')"><div style="display:flex; flex-direction:column; gap:4px;"><span class="dd-text" [style.color]="activeColorSet === 'Default' ? '#1a73e8' : '#000'">Default</span><div style="display:flex; gap:2px;"><div style="width:12px; height:12px; background:#1e5e3c;"></div><div style="width:12px; height:12px; background:#4285f4;"></div><div style="width:12px; height:12px; background:#fbbc04;"></div><div style="width:12px; height:12px; background:#34a853;"></div><div style="width:12px; height:12px; background:#9aa0a6;"></div><div style="width:12px; height:12px; background:#ff6d00;"></div><div style="width:12px; height:12px; background:#00bcd4;"></div></div></div></div>
                      <div class="dd-item" (click)="applyColorSet('Pinstripes')"><div style="display:flex; flex-direction:column; gap:4px;"><span class="dd-text" [style.color]="activeColorSet === 'Pinstripes' ? '#1a73e8' : '#000'">Pinstripes</span><div style="display:flex; gap:2px;"><div style="width:12px; height:12px; background:#5a4342;"></div><div style="width:12px; height:12px; background:#f4f1e1;"></div><div style="width:12px; height:12px; background:#4caf50;"></div><div style="width:12px; height:12px; background:#cddc39;"></div><div style="width:12px; height:12px; background:#00bcd4;"></div><div style="width:12px; height:12px; background:#ff9800;"></div><div style="width:12px; height:12px; background:#795548;"></div></div></div></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Add Color Set')"><span class="material-symbols-outlined dd-icon" style="color:#34a853;">add_circle</span><span class="dd-text">Add Color Set</span></div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>
                  
                  <div class="dd-item" (click)="importDesign()"><span class="material-symbols-outlined dd-icon">import_contacts</span><span class="dd-text">Import Design...</span></div>
                  <div class="dd-item" (click)="setPageBorders()"><span class="material-symbols-outlined dd-icon">border_outer</span><span class="dd-text">Page Borders</span></div>
                  <div class="dd-item" (click)="setPageBackground()"><span class="material-symbols-outlined dd-icon">format_color_fill</span><span class="dd-text">Page Background...</span></div>
                  
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="showToast('More Design Options')"><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More Design Options...</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('pagesetup', $event)" [class.active]="activeMenu === 'pagesetup'">
                Page Setup
                <div class="dropdown" *ngIf="activeMenu === 'pagesetup'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subPageSize)">
                    <span class="material-symbols-outlined dd-icon">crop_portrait</span><span class="dd-text">Page Size: <span style="color:#1a73e8;">{{ activePageSize }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subPageSize class="sub-dropdown" style="max-height: 350px; overflow-y: auto;">
                      <div class="dd-item" (click)="setPageSize('A4')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'A4' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'A4' ? '#1a73e8' : '#000'">A4 (8.27" X 11.69")</span></div>
                      <div class="dd-item" (click)="setPageSize('Letter')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Letter' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Letter' ? '#1a73e8' : '#000'">Letter (8.5" X 11")</span></div>
                      <div class="dd-item" (click)="setPageSize('Legal')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Legal' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Legal' ? '#1a73e8' : '#000'">Legal (8.5" X 14")</span></div>
                      <div class="dd-item" (click)="setPageSize('Executive')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Executive' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Executive' ? '#1a73e8' : '#000'">Executive (7.25" X 10.5")</span></div>
                      <div class="dd-item" (click)="setPageSize('Envelope_10')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Envelope_10' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Envelope_10' ? '#1a73e8' : '#000'">Envelope_10 (4.13" X 9.5")</span></div>
                      <div class="dd-item" (click)="setPageSize('Tabloid')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Tabloid' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Tabloid' ? '#1a73e8' : '#000'">Tabloid (11" X 17")</span></div>
                      <div class="dd-item" (click)="setPageSize('Statement')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Statement' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Statement' ? '#1a73e8' : '#000'">Statement (5.5" X 8.5")</span></div>
                      <div class="dd-item" (click)="setPageSize('Folio')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'Folio' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'Folio' ? '#1a73e8' : '#000'">Folio (8" X 13")</span></div>
                      <div class="dd-item" (click)="setPageSize('A3')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'A3' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'A3' ? '#1a73e8' : '#000'">A3 (11.69" X 16.54")</span></div>
                      <div class="dd-item" (click)="setPageSize('A5')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'A5' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'A5' ? '#1a73e8' : '#000'">A5 (5.83" X 8.27")</span></div>
                      <div class="dd-item" (click)="setPageSize('B4')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'B4' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'B4' ? '#1a73e8' : '#000'">B4 (10.12" X 14.33")</span></div>
                      <div class="dd-item" (click)="setPageSize('B5')"><span class="material-symbols-outlined dd-icon" [style.color]="activePageSize === 'B5' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activePageSize === 'B5' ? '#1a73e8' : '#000'">B5 (7.17" X 10.12")</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('More Sizes')"><span class="dd-text">More</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subOrientation)">
                    <span class="material-symbols-outlined dd-icon">description</span><span class="dd-text">Orientation: <span style="color:#1a73e8;">{{ activeOrientation }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subOrientation class="sub-dropdown">
                      <div class="dd-item" (click)="setPageOrientation('Portrait')"><span class="material-symbols-outlined dd-icon" [style.color]="activeOrientation === 'Portrait' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeOrientation === 'Portrait' ? '#1a73e8' : '#000'">description</span><span class="dd-text" [style.color]="activeOrientation === 'Portrait' ? '#1a73e8' : '#000'">Portrait</span></div>
                      <div class="dd-item" (click)="setPageOrientation('Landscape')"><span class="material-symbols-outlined dd-icon" [style.color]="activeOrientation === 'Landscape' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeOrientation === 'Landscape' ? '#1a73e8' : '#000'">note</span><span class="dd-text" [style.color]="activeOrientation === 'Landscape' ? '#1a73e8' : '#000'">Landscape</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subColumns)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">view_column</span><span class="dd-text" style="color:#1a73e8;">Columns: <span style="color:#1a73e8;">{{ activeColumns }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subColumns class="sub-dropdown">
                      <div class="dd-item" (click)="setPageColumns('One')"><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'One' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'One' ? '#1a73e8' : '#000'">view_column</span><span class="dd-text" [style.color]="activeColumns === 'One' ? '#1a73e8' : '#000'">One</span></div>
                      <div class="dd-item" (click)="setPageColumns('Two')"><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Two' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Two' ? '#1a73e8' : '#000'">view_column_2</span><span class="dd-text" [style.color]="activeColumns === 'Two' ? '#1a73e8' : '#000'">Two</span></div>
                      <div class="dd-item" (click)="setPageColumns('Three')"><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Three' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Three' ? '#1a73e8' : '#000'">view_week</span><span class="dd-text" [style.color]="activeColumns === 'Three' ? '#1a73e8' : '#000'">Three</span></div>
                      <div class="dd-item" (click)="setPageColumns('Left')"><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Left' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Left' ? '#1a73e8' : '#000'">format_align_left</span><span class="dd-text" [style.color]="activeColumns === 'Left' ? '#1a73e8' : '#000'">Left</span></div>
                      <div class="dd-item" (click)="setPageColumns('Right')"><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Right' ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon" [style.color]="activeColumns === 'Right' ? '#1a73e8' : '#000'">format_align_right</span><span class="dd-text" [style.color]="activeColumns === 'Right' ? '#1a73e8' : '#000'">Right</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('More Columns')"><span class="dd-text">More</span></div>
                    </div>
                  </div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMargins)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">border_clear</span><span class="dd-text" style="color:#1a73e8;">Margins: <span style="color:#1a73e8;">{{ activeMargins }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subMargins class="sub-dropdown">
                      <div class="dd-item" (click)="setPageMargins('Normal')">
                        <span class="material-symbols-outlined dd-icon" [style.color]="activeMargins === 'Normal' ? '#1a73e8' : 'transparent'">check</span>
                        <div [style.borderColor]="activeMargins === 'Normal' ? '#1a73e8' : '#dadce0'" style="border: 1px solid; width: 24px; height: 32px; display: flex; align-items: center; justify-content: center; margin-right: 8px;"><div [style.borderColor]="activeMargins === 'Normal' ? '#1a73e8' : '#dadce0'" style="border: 1px dashed; width: 14px; height: 22px;"></div></div>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeMargins === 'Normal' ? '#1a73e8' : '#000'">Normal</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">Top : 1in, Left : 1in<br>Bottom: 1in, Right: 1in</span></div>
                      </div>
                      <div class="dd-item" (click)="setPageMargins('Narrow')">
                        <span class="material-symbols-outlined dd-icon" [style.color]="activeMargins === 'Narrow' ? '#1a73e8' : 'transparent'">check</span>
                        <div [style.borderColor]="activeMargins === 'Narrow' ? '#1a73e8' : '#dadce0'" style="border: 1px solid; width: 24px; height: 32px; display: flex; align-items: center; justify-content: center; margin-right: 8px;"><div [style.borderColor]="activeMargins === 'Narrow' ? '#1a73e8' : '#dadce0'" style="border: 1px dashed; width: 18px; height: 26px;"></div></div>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeMargins === 'Narrow' ? '#1a73e8' : '#000'">Narrow</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">Top : 0.5in, Left : 0.5in<br>Bottom: 0.5in, Right: 0.5in</span></div>
                      </div>
                      <div class="dd-item" (click)="setPageMargins('Moderate')">
                        <span class="material-symbols-outlined dd-icon" [style.color]="activeMargins === 'Moderate' ? '#1a73e8' : 'transparent'">check</span>
                        <div [style.borderColor]="activeMargins === 'Moderate' ? '#1a73e8' : '#dadce0'" style="border: 1px solid; width: 24px; height: 32px; display: flex; align-items: center; justify-content: center; margin-right: 8px;"><div [style.borderColor]="activeMargins === 'Moderate' ? '#1a73e8' : '#dadce0'" style="border: 1px dashed; width: 16px; height: 22px;"></div></div>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeMargins === 'Moderate' ? '#1a73e8' : '#000'">Moderate</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">Top : 1in, Left : 0.75in<br>Bottom: 1in, Right: 0.75in</span></div>
                      </div>
                      <div class="dd-item" (click)="setPageMargins('Wide')">
                        <span class="material-symbols-outlined dd-icon" [style.color]="activeMargins === 'Wide' ? '#1a73e8' : 'transparent'">check</span>
                        <div [style.borderColor]="activeMargins === 'Wide' ? '#1a73e8' : '#dadce0'" style="border: 1px solid; width: 24px; height: 32px; display: flex; align-items: center; justify-content: center; margin-right: 8px;"><div [style.borderColor]="activeMargins === 'Wide' ? '#1a73e8' : '#dadce0'" style="border: 1px dashed; width: 10px; height: 22px;"></div></div>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeMargins === 'Wide' ? '#1a73e8' : '#000'">Wide</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">Top : 1in, Left : 2.28in<br>Bottom: 1in, Right: 2.28in</span></div>
                      </div>
                      <div class="dd-item" (click)="setPageMargins('None')">
                        <span class="material-symbols-outlined dd-icon" [style.color]="activeMargins === 'None' ? '#1a73e8' : 'transparent'">check</span>
                        <div [style.borderColor]="activeMargins === 'None' ? '#1a73e8' : '#dadce0'" style="border: 1px solid; width: 24px; height: 32px; margin-right: 8px;"></div>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text" [style.color]="activeMargins === 'None' ? '#1a73e8' : '#000'">None</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">Top : 0in, Left : 0in<br>Bottom: 0in, Right: 0in</span></div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Custom Margins')"><span class="dd-text" style="font-weight: 500;">Custom margins ...</span></div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subAdvancedPageSetup)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">settings_applications</span><span class="dd-text" style="color:#1a73e8;">Advanced Page Setup</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subAdvancedPageSetup class="sub-dropdown" style="min-width: 250px; padding: 0;">
                      <div [style.background]="activeAdvancedSetup === 'Document-Level Setup' ? '#e8f0fe' : '#fff'" style="padding: 12px; cursor: pointer;" (click)="setAdvancedSetup('Document-Level Setup')">
                        <div [style.color]="activeAdvancedSetup === 'Document-Level Setup' ? '#1a73e8' : '#000'" style="font-weight: 600; margin-bottom: 4px;">Document-Level Setup</div>
                        <div style="font-size: 11px; color: #5f6368; line-height: 1.4;">Configure the settings that you wish to apply to the entire document.</div>
                      </div>
                      <div [style.background]="activeAdvancedSetup === 'Section-Level Setup' ? '#e8f0fe' : '#fff'" style="padding: 12px; cursor: pointer;" (click)="setAdvancedSetup('Section-Level Setup')">
                        <div [style.color]="activeAdvancedSetup === 'Section-Level Setup' ? '#1a73e8' : '#000'" style="font-weight: 600; margin-bottom: 4px;">Section-Level Setup</div>
                        <div style="font-size: 11px; color: #5f6368; line-height: 1.4;">Configure the settings that you wish to apply only to a section.</div>
                      </div>
                    </div>
                  </div>

                  <div class="dd-sep"></div>

                  <div class="dd-item" (click)="showToast('More Page Setup Options')"><span class="material-symbols-outlined dd-icon">more_horiz</span><span class="dd-text">More Page Setup Options</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('review', $event)" [class.active]="activeMenu === 'review'">
                Review
                <div class="dropdown" *ngIf="activeMenu === 'review'">
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertComment()"><span class="material-symbols-outlined dd-icon">add_comment</span><span class="dd-text">Add Comments</span></div>
                  <div class="dd-item" (click)="showToast('Show Comments')"><span class="material-symbols-outlined dd-icon">chat</span><span class="dd-text">Show Comments</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subCollab)">
                    <span class="material-symbols-outlined dd-icon" style="color:#ccc;">group</span><span class="dd-text" style="color:#ccc;">Collaboration : <span style="color:#8ab4f8;">Off</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#ccc;">chevron_right</span>
                    <div #subCollab class="sub-dropdown"><div class="dd-item" style="color:#ccc; cursor:default;"><span class="dd-text">Off</span></div></div>
                  </div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subTrack)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">published_with_changes</span><span class="dd-text" style="color:#1a73e8;">Track Changes: <span style="color:#1a73e8;">{{ activeTrackChanges }}</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subTrack class="sub-dropdown">
                      <div class="dd-item" (click)="setTrackChanges('On')"><span class="material-symbols-outlined dd-icon" [style.color]="activeTrackChanges === 'On' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeTrackChanges === 'On' ? '#1a73e8' : '#000'">On</span></div>
                      <div class="dd-item" (click)="setTrackChanges('Off')"><span class="material-symbols-outlined dd-icon" [style.color]="activeTrackChanges === 'Off' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeTrackChanges === 'Off' ? '#1a73e8' : '#000'">Off</span></div>
                    </div>
                  </div>
                  <div class="dd-item" (click)="showToast('View Suggestions')"><span class="material-symbols-outlined dd-icon">rule</span><span class="dd-text">View Suggestions</span></div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMarkup)">
                    <span class="material-symbols-outlined dd-icon" style="color:#ccc;">find_in_page</span><span class="dd-text" style="color:#ccc;">Markup View : <span style="color:#8ab4f8;">All Markup</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#ccc;">chevron_right</span>
                    <div #subMarkup class="sub-dropdown"><div class="dd-item" style="color:#ccc; cursor:default;"><span class="dd-text">All Markup</span></div></div>
                  </div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subMarkupColor)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">person_search</span><span class="dd-text" style="color:#1a73e8;">Markup Color</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subMarkupColor class="sub-dropdown" style="min-width: 300px;">
                      <div style="padding: 12px;">
                         <div style="color: #1a73e8; font-weight: 600; font-size: 13px; margin-bottom: 4px;">Pick Your Color</div>
                         <div style="font-size: 11px; color: #5f6368; line-height: 1.4; margin-bottom: 12px;">Your cursor and the changes you make will appear in this color in this document.</div>
                         <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px;">
                            <div *ngFor="let color of markupColors" (click)="setMarkupColor(color)" [style.background]="color" style="width: 20px; height: 20px; border-radius: 50%; display: flex; align-items: center; justify-content: center; cursor: pointer;">
                               <span *ngIf="activeMarkupColor === color" class="material-symbols-outlined" style="color: #fff; font-size: 14px;">check</span>
                            </div>
                         </div>
                         <div style="display: flex; align-items: center; gap: 8px;" (click)="toggleMarkupDefault()">
                            <span class="material-symbols-outlined" style="font-size: 16px; color: #1a73e8; cursor: pointer;">{{ markupDefault ? 'check_box' : 'check_box_outline_blank' }}</span>
                            <span style="font-size: 12px; color: #1a73e8; cursor: pointer;">Use as default color for new documents.</span>
                         </div>
                      </div>
                    </div>
                  </div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="compareVersions()"><span class="material-symbols-outlined dd-icon">compare</span><span class="dd-text">Compare Versions</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="combineRevisions()"><span class="material-symbols-outlined dd-icon">merge_type</span><span class="dd-text">Combine Revisions</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="lockUnlockContent()"><span class="material-symbols-outlined dd-icon">lock</span><span class="dd-text">Lock/Unlock Content</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="maskContent()"><span class="material-symbols-outlined dd-icon">visibility_off</span><span class="dd-text">Mask/Unmask Content</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="toggleNotificationSettings()">
                     <span class="material-symbols-outlined dd-icon" [style.color]="activeNotifications ? '#1a73e8' : 'transparent'">check</span>
                     <span class="material-symbols-outlined dd-icon" style="margin-left: 0;">mark_email_unread</span>
                     <span class="dd-text">Notification Settings</span>
                  </div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('tools', $event)" [class.active]="activeMenu === 'tools'">
                Tools
                <div class="dropdown" *ngIf="activeMenu === 'tools'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subSpellCheck)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">spellcheck</span><span class="dd-text" style="color:#1a73e8;">Spell Check</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subSpellCheck class="sub-dropdown">
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Writing Suggestions</div>
                      <div class="dd-item" (click)="toggleSpellSetting('Spelling Errors')"><span class="material-symbols-outlined dd-icon" [style.color]="activeSpellErrors ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon">spellcheck</span><span class="dd-text">Spelling Errors</span></div>
                      <div class="dd-item" (click)="toggleSpellSetting('Grammar')"><span class="material-symbols-outlined dd-icon" [style.color]="activeGrammar ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon">history_edu</span><span class="dd-text">Grammar</span></div>
                      <div class="dd-item" (click)="toggleSpellSetting('Writing Quality')"><span class="material-symbols-outlined dd-icon" [style.color]="activeWritingQuality ? '#1a73e8' : 'transparent'">check</span><span class="material-symbols-outlined dd-icon">draw</span><span class="dd-text">Writing Quality</span></div>
                      <div class="dd-sep"></div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Language</div>
                      <div class="dd-item" (click)="showToast('Language')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">check</span><span class="material-symbols-outlined dd-icon">language</span><span class="dd-text">Language: <span style="color:#1a73e8;">{{ activeLanguage }}</span></span></div>
                    </div>
                  </div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" style="color:#ccc; cursor:default;"><span class="material-symbols-outlined dd-icon" style="color:#ccc;">table_chart</span><span class="dd-text">Text to Table</span></div>
                  <div class="dd-item has-sub" (mousedown)="$event.preventDefault()" (click)="activeTransliteration !== 'None' ? setTransliteration(activeTransliteration) : null" (mouseenter)="positionSubmenu($event, subTransliteration)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">translate</span><span class="dd-text" style="color:#1a73e8;">Quick Translate Selection To: <span style="color:#1a73e8;">{{ activeTransliteration }}</span> <span class="material-symbols-outlined" style="font-size:14px; margin-left:2px; color:#5f6368;">info</span></span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subTransliteration class="sub-dropdown" style="max-height: 350px; overflow-y: auto;">
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="setTransliteration('None')"><span class="material-symbols-outlined dd-icon" [style.color]="activeTransliteration === 'None' ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeTransliteration === 'None' ? '#1a73e8' : '#000'">None</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" *ngFor="let lang of ['Bengali','French','Gujarati','Hindi','Kannada','Malayalam','Marathi','Odia','Punjabi','Spanish','Tamil','Telugu']" (click)="setTransliteration(lang)"><span class="material-symbols-outlined dd-icon" [style.color]="activeTransliteration === lang ? '#1a73e8' : 'transparent'">check</span><span class="dd-text" [style.color]="activeTransliteration === lang ? '#1a73e8' : '#000'">{{lang}}</span></div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (click)="showToast('Request Language')"><span class="dd-text" style="color:#1a73e8;">Request Language</span></div>
                    </div>
                  </div>
                  <div class="dd-item" (click)="showToast('Focus Typing')"><span class="material-symbols-outlined dd-icon">center_focus_strong</span><span class="dd-text">Focus Typing</span></div>
                  <div class="dd-item" (click)="showToast('Typewriter Sound')"><span class="material-symbols-outlined dd-icon">keyboard</span><span class="dd-text">Typewriter Sound</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="openDictionary()"><span class="material-symbols-outlined dd-icon">auto_stories</span><span class="dd-text">Thesaurus</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="activeAutocorrect = !activeAutocorrect"><span class="material-symbols-outlined dd-icon" [style.color]="activeAutocorrect ? '#1a73e8' : 'transparent'">check</span><span class="dd-text">Autocorrect</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Personal Dictionary is currently empty.')"><span class="material-symbols-outlined dd-icon">import_contacts</span><span class="dd-text">Personal Dictionary</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showWordCount()"><span class="material-symbols-outlined dd-icon">pin</span><span class="dd-text">Word Count</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('No images found in the current selection.')"><span class="material-symbols-outlined dd-icon">image</span><span class="dd-text">View Document Images</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subExtensions)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">extension</span><span class="dd-text" style="color:#1a73e8;">Extensions</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subExtensions class="sub-dropdown" style="max-height: 400px; overflow-y: auto;">
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Utility</div>
                      <div class="dd-item" (click)="showToast('WordPress')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">public</span><span class="dd-text">WordPress</span></div>
                      <div class="dd-item" (click)="showToast('Signeasy')"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">draw</span><span class="dd-text">Signeasy</span></div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Publishing</div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('WordPress.org linked')"><span class="material-symbols-outlined dd-icon">public</span><span class="dd-text">WordPress.org</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('VMail Learn activated')"><span class="material-symbols-outlined dd-icon" style="color:#d32f2f;">school</span><span class="dd-text">VMail Learn</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Blogger linked')"><span class="material-symbols-outlined dd-icon" style="color:#f57c00;">rss_feed</span><span class="dd-text">Blogger</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('VMail Connect activated')"><span class="material-symbols-outlined dd-icon" style="color:#1976d2;">hub</span><span class="dd-text">VMail Connect</span></div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">AI Assistance</div>
                      <div class="dd-item" (click)="showToast('ChatGPT Assistant')"><span class="material-symbols-outlined dd-icon" style="color:#388e3c;">smart_toy</span><span class="dd-text">ChatGPT Assistant</span></div>
                      <div class="dd-item" (click)="showToast('Cohere Assistant')"><span class="material-symbols-outlined dd-icon" style="color:#5e35b1;">smart_toy</span><span class="dd-text">Cohere Assistant</span></div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Diagramming</div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('VMail ChemStudio loaded')"><span class="material-symbols-outlined dd-icon" style="color:#d32f2f;">science</span><span class="dd-text">VMail ChemStudio</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Mermaid Chart loaded')"><span class="material-symbols-outlined dd-icon" style="color:#c2185b;">schema</span><span class="dd-text">Mermaid Chart</span></div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Asset Library</div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Unsplash gallery open')"><span class="material-symbols-outlined dd-icon">image</span><span class="dd-text">Unsplash</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('VMail Analytics loaded')"><span class="material-symbols-outlined dd-icon" style="color:#d32f2f;">analytics</span><span class="dd-text">VMail Analytics</span></div>
                    </div>
                  </div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="engagementModalOpen = true; closeMenus()"><span class="material-symbols-outlined dd-icon">monitoring</span><span class="dd-text">Engagement Insights</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('fields', $event)" [class.active]="activeMenu === 'fields'">
                Fields
                <div class="dropdown" *ngIf="activeMenu === 'fields'">
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subDate)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">calendar_today</span><span class="dd-text" style="color:#1a73e8;">Date</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subDate class="sub-dropdown" style="min-width: 200px;">
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Static Date</div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(getTodayDate())">
                        <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">check</span>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text">Today's Date</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">({{ getTodayDate() }})</span></div>
                      </div>
                      <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Dynamic Dates</div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(getTodayDate())">
                        <span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text">Current Date</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">({{ getTodayDate() }})</span></div>
                      </div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('06/17/2026')">
                        <span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text">Created Date</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">(06/17/2026)</span></div>
                      </div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('06/19/2026')">
                        <span class="material-symbols-outlined dd-icon" style="color:transparent;">check</span>
                        <div style="display:flex; flex-direction:column;"><span class="dd-text">Last Edited Date</span><span class="dd-text" style="font-size: 11px; color:#5f6368;">(06/19/2026)</span></div>
                      </div>
                      <div class="dd-sep"></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Custom Date Field]')"><span class="material-symbols-outlined dd-icon">edit_calendar</span><span class="dd-text">Create custom date field...</span></div>
                    </div>
                  </div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subPageNumber)">
                    <span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">123</span><span class="dd-text" style="color:#1a73e8;">Page Number</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                    <div #subPageNumber class="sub-dropdown">
                      <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subInsertPageNumber)">
                        <span class="dd-text" style="color:#1a73e8;">Insert Page Number</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#1a73e8;">chevron_right</span>
                        <div #subInsertPageNumber class="sub-dropdown">
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Top Left)]')"><span class="dd-text">Top Left</span></div>
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Top Center)]')"><span class="dd-text">Top Center</span></div>
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Top Right)]')"><span class="dd-text">Top Right</span></div>
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Bottom Left)]')"><span class="dd-text">Bottom Left</span></div>
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Bottom Center)]')"><span class="dd-text">Bottom Center</span></div>
                           <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('[Page Number (Bottom Right)]')"><span class="dd-text">Bottom Right</span></div>
                        </div>
                      </div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Page Number Formatting is disabled')"><span class="dd-text">Format Page Number</span></div>
                      <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="showToast('Page numbers removed')"><span class="dd-text">Remove Page Numbers</span></div>
                    </div>
                  </div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(pageCountArray.length.toString())"><span class="material-symbols-outlined dd-icon">pin</span><span class="dd-text">Page Count</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(auth.user?.name || 'Author Name')"><span class="material-symbols-outlined dd-icon">person_outline</span><span class="dd-text">Author Name</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(title || 'Untitled Document')"><span class="material-symbols-outlined dd-icon">description</span><span class="dd-text">Document Name</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField('v1.0')"><span class="material-symbols-outlined dd-icon">history</span><span class="dd-text">Document Version</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(auth.user ? (auth.user.name || '').split(' ')[0] : '[First Name]')"><span class="material-symbols-outlined dd-icon">badge</span><span class="dd-text">First Name</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(auth.user ? (auth.user.name || '').split(' ').slice(1).join(' ') : '[Last Name]')"><span class="material-symbols-outlined dd-icon">badge</span><span class="dd-text">Last Name</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertField(auth.user?.email || '[Email]')"><span class="material-symbols-outlined dd-icon">mail</span><span class="dd-text">Email</span></div>
                  <div class="dd-item" (mousedown)="$event.preventDefault()" (click)="insertPhoneField()"><span class="material-symbols-outlined dd-icon">call</span><span class="dd-text">Phone</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('automate', $event)" [class.active]="activeMenu === 'automate'">
                Automate
                <div class="dropdown" *ngIf="activeMenu === 'automate'" style="width: 350px;">
                  <div style="padding: 8px 12px; font-weight: 600; font-size: 11px; color: #000;">Convert to</div>
                  <div class="dd-item" (click)="toggleAutomateSidebar('Merge Template')" style="align-items: flex-start; padding: 12px; height: auto;">
                    <span class="material-symbols-outlined dd-icon" style="color:#f57c00; margin-top:2px;">merge_type</span>
                    <div style="display:flex; flex-direction:column;">
                      <span class="dd-text" style="white-space: normal;">Merge Template</span>
                      <span class="dd-text" style="font-size: 11px; color:#5f6368; white-space: normal; line-height: 1.4; margin-top:4px;">Use this to generate personalized documents in bulk that can be emailed, downloaded, or sent for signature.</span>
                    </div>
                  </div>
                  <div class="dd-item" (click)="toggleAutomateSidebar('Fillable Template')" style="align-items: flex-start; padding: 12px; height: auto;">
                    <span class="material-symbols-outlined dd-icon" style="color:#7b1fa2; margin-top:2px;">dynamic_form</span>
                    <div style="display:flex; flex-direction:column;">
                      <span class="dd-text" style="white-space: normal;">Fillable Template</span>
                      <span class="dd-text" style="font-size: 11px; color:#5f6368; white-space: normal; line-height: 1.4; margin-top:4px;">Use this to create and publish forms that allow users to fill and you to collect responses.</span>
                    </div>
                  </div>
                  <div class="dd-item" (click)="toggleAutomateSidebar('Sign Template')" style="align-items: flex-start; padding: 12px; height: auto;">
                    <span class="material-symbols-outlined dd-icon" style="color:#388e3c; margin-top:2px;">draw</span>
                    <div style="display:flex; flex-direction:column;">
                      <span class="dd-text" style="white-space: normal;">Sign Template</span>
                      <span class="dd-text" style="font-size: 11px; color:#5f6368; white-space: normal; line-height: 1.4; margin-top:4px;">Use this to collect signatures by automating document delivery.</span>
                    </div>
                  </div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('help', $event)" [class.active]="activeMenu === 'help'">
                Help
                <div class="dropdown" *ngIf="activeMenu === 'help'">
                  <div class="dd-item" (click)="showToast('Whats New')" style="background: #e8f0fe;"><span class="material-symbols-outlined dd-icon" style="color:#1a73e8;">redeem</span><span class="dd-text" style="color:#1a73e8;">What's New</span></div>
                  <div class="dd-item" (click)="showToast('User Guide')"><span class="material-symbols-outlined dd-icon">help_center</span><span class="dd-text">User Guide</span></div>
                  <div class="dd-item" (click)="showToast('Knowledge Base')"><span class="material-symbols-outlined dd-icon">lightbulb</span><span class="dd-text">Knowledge Base</span></div>
                  <div class="dd-item" (click)="showToast('Community')"><span class="material-symbols-outlined dd-icon">local_library</span><span class="dd-text">Community</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="showToast('Accessibility options')"><span class="material-symbols-outlined dd-icon">accessibility_new</span><span class="dd-text">Accessibility options ...</span></div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subReadAloud)">
                    <span class="material-symbols-outlined dd-icon">record_voice_over</span><span class="dd-text">Read aloud to screen reader</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subReadAloud class="sub-dropdown"><div class="dd-item"><span class="dd-text">Play</span></div><div class="dd-item"><span class="dd-text">Pause</span></div></div>
                  </div>
                  <div class="dd-item has-sub" (mouseenter)="positionSubmenu($event, subNavigateShortcuts)">
                    <span class="material-symbols-outlined dd-icon">account_tree</span><span class="dd-text">Navigate with shortcuts</span><span class="material-symbols-outlined" style="margin-left:auto; font-size:16px; color:#9aa0a6;">chevron_right</span>
                    <div #subNavigateShortcuts class="sub-dropdown"><div class="dd-item"><span class="dd-text">Next Heading</span></div><div class="dd-item"><span class="dd-text">Previous Heading</span></div></div>
                  </div>
                  <div class="dd-item" (click)="showToast('Keyboard Shortcuts')"><span class="material-symbols-outlined dd-icon">keyboard</span><span class="dd-text">Keyboard Shortcuts ...</span></div>
                </div>
            </div>
          </div>
        </div>

      <!-- Formatting Toolbar -->
      <div class="fmt-bar" (mousedown)="$event.preventDefault()">
        <button class="fb" (click)="exec('undo')" title="Undo"><span class="material-symbols-outlined">undo</span></button>
        <button class="fb" (click)="exec('redo')" title="Redo"><span class="material-symbols-outlined">redo</span></button>
        <button class="fb" (click)="printDoc()" title="Print"><span class="material-symbols-outlined">print</span></button>
        <span class="sep"></span>

        <div class="menu-item style-dropdown" (click)="toggleMenu('style', $event)" [class.active]="activeMenu === 'style'" title="Paragraph style">
          <span style="max-width:75px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; display:inline-block; vertical-align:middle;">{{ activeBlockStyle }}</span> <span class="material-symbols-outlined arrow-icon" style="vertical-align:middle;">expand_more</span>
          <div class="dropdown style-dd" *ngIf="activeMenu === 'style'">
            <div class="dd-item" (click)="changeStyle('p', 'Normal')"><span class="dd-text" style="font-size:14px;">Normal</span></div>
            <div class="dd-item" (click)="changeStyle('h1', 'Title')"><span class="dd-text" style="font-size:22px;font-weight:700;">Title</span></div>
            <div class="dd-item" (click)="changeStyle('h2', 'Subtitle')"><span class="dd-text" style="font-size:16px;color:#5f6368;">Subtitle</span></div>
            <div class="dd-item" (click)="changeStyle('h3', 'Heading 1')"><span class="dd-text" style="font-size:18px;font-weight:600;">Heading 1</span></div>
            <div class="dd-item" (click)="changeStyle('h4', 'Heading 2')"><span class="dd-text" style="font-size:15px;font-weight:600;">Heading 2</span></div>
            <div class="dd-item" (click)="changeStyle('h5', 'Heading 3')"><span class="dd-text" style="font-size:13px;font-weight:600;">Heading 3</span></div>
          </div>
        </div>
        
        <span class="sep"></span>

        <div class="menu-item font-dropdown" (click)="toggleMenu('font', $event)" [class.active]="activeMenu === 'font'" title="Font family">
          <span [style.font-family]="activeFontFamily" style="max-width:80px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; display:inline-block; vertical-align:middle;">{{ activeFontFamily }}</span> <span class="material-symbols-outlined arrow-icon" style="vertical-align:middle;">expand_more</span>
          <div class="dropdown font-dd" *ngIf="activeMenu === 'font'">
            <div class="dd-item" *ngFor="let font of fonts" (click)="changeFont(font)">
              <span class="dd-text" [style.font-family]="font">{{ font }}</span>
            </div>
          </div>
        </div>

        <span class="sep"></span>

        <div class="font-size-control" style="position: relative;">
          <button class="fb size-btn" (click)="decrementFontSize()" title="Decrease font size"><span class="material-symbols-outlined">remove</span></button>
          
          <div class="size-input-wrapper" (click)="showFontSizeMenu = !showFontSizeMenu; $event.stopPropagation()">
            <input class="size-input" [(ngModel)]="activeFontSize" (change)="onFontSizeInputChange()" (click)="showFontSizeMenu = !showFontSizeMenu; $event.stopPropagation()" />
            <span class="material-symbols-outlined">arrow_drop_down</span>
          </div>

          <button class="fb size-btn" (click)="incrementFontSize()" title="Increase font size"><span class="material-symbols-outlined">add</span></button>
          <div class="dropdown" *ngIf="showFontSizeMenu" (click)="$event.stopPropagation()" style="position: absolute; top: 100%; left: 24px; width: 52px; min-width: 52px; z-index: 1000; max-height: 250px; overflow-y: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1); background: white; border: 1px solid #ccc; padding: 4px 0; display: block; border-radius: 4px;">
            <div class="dd-item" style="padding: 6px 12px; text-align: center; cursor: pointer; border-bottom: none;" *ngFor="let s of [8,9,10,11,12,14,18,24,30,36,48,60,72]" (click)="changeFontSize(s)">{{ s }}</div>
          </div>
        </div>

        <span class="sep"></span>

        <button class="fb" (click)="exec('bold')" [class.active-fb]="isBold" title="Bold (Ctrl+B)"><span class="material-symbols-outlined">format_bold</span></button>
        <button class="fb" (click)="exec('italic')" [class.active-fb]="isItalic" title="Italic (Ctrl+I)"><span class="material-symbols-outlined">format_italic</span></button>
        <button class="fb" (click)="exec('underline')" [class.active-fb]="isUnderline" title="Underline (Ctrl+U)"><span class="material-symbols-outlined">format_underlined</span></button>
        <button class="fb" (click)="exec('strikeThrough')" [class.active-fb]="isStrikethrough" title="Strikethrough"><span class="material-symbols-outlined">strikethrough_s</span></button>
        
        <div class="menu-item" (click)="toggleMenu('textcolor', $event)" [class.active]="activeMenu === 'textcolor'" style="padding: 0 6px; position:relative; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:0px; height: 28px;" title="Text color">
           <span class="material-symbols-outlined" style="font-size: 20px; margin-bottom: 2px;">format_color_text</span>
           <span class="color-indicator" [style.background]="activeColor || '#000000'"></span>
           <div class="dropdown" *ngIf="activeMenu === 'textcolor'" style="min-width: 200px; padding:12px; cursor:default;" (click)="$event.stopPropagation()">
              <div style="font-size:12px; font-weight:600; color:#5f6368; margin-bottom:8px;">Theme Colors</div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="execVal('foreColor', c); activeColor=c; closeMenus()"></div></div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="execVal('foreColor', c); activeColor=c; closeMenus()"></div></div>
              <div style="font-size:12px; font-weight:600; color:#5f6368; margin:12px 0 8px 0;">Standard Colors</div>
              <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="execVal('foreColor', c); activeColor=c; closeMenus()"></div></div>
           </div>
        </div>
        
        <div class="menu-item" (click)="toggleMenu('bgcolor', $event)" [class.active]="activeMenu === 'bgcolor'" style="padding: 0 6px; position:relative; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:0px; height: 28px;" title="Highlight color">
           <span class="material-symbols-outlined" style="font-size: 20px; margin-bottom: 2px;">format_ink_highlighter</span>
           <span class="color-indicator" [style.background]="activeHighlight || 'transparent'"></span>
           <div class="dropdown" *ngIf="activeMenu === 'bgcolor'" style="min-width: 200px; padding:12px; cursor:default;" (click)="$event.stopPropagation()">
              <div class="dd-item" (click)="execVal('hiliteColor', 'transparent'); activeHighlight='transparent'; closeMenus()" style="margin:0 -12px 8px -12px; border-bottom:1px solid #eee; padding-bottom:8px;"><span class="material-symbols-outlined dd-icon">format_color_reset</span><span class="dd-text">No Fill</span></div>
              <div style="font-size:12px; font-weight:600; color:#5f6368; margin-bottom:8px;">Theme Colors</div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="execVal('hiliteColor', c); activeHighlight=c; closeMenus()"></div></div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="execVal('hiliteColor', c); activeHighlight=c; closeMenus()"></div></div>
              <div style="font-size:12px; font-weight:600; color:#5f6368; margin:12px 0 8px 0;">Standard Colors</div>
              <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="execVal('hiliteColor', c); activeHighlight=c; closeMenus()"></div></div>
           </div>
        </div>

        <span class="sep"></span>

        <div class="menu-item" (click)="toggleMenu('align', $event)" [class.active]="activeMenu === 'align'" title="Align">
          <span class="material-symbols-outlined" style="font-size: 20px;">format_align_left</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'align'" style="min-width: 170px;">
            <div class="dd-item" (click)="exec('justifyLeft')"><span class="material-symbols-outlined dd-icon">format_align_left</span><span class="dd-text">Align Left</span><span class="dd-hint">Ctrl+Shift+L</span></div>
            <div class="dd-item" (click)="exec('justifyCenter')"><span class="material-symbols-outlined dd-icon">format_align_center</span><span class="dd-text">Align Center</span><span class="dd-hint">Ctrl+Shift+E</span></div>
            <div class="dd-item" (click)="exec('justifyRight')"><span class="material-symbols-outlined dd-icon">format_align_right</span><span class="dd-text">Align Right</span><span class="dd-hint">Ctrl+Shift+R</span></div>
            <div class="dd-item" (click)="exec('justifyFull')"><span class="material-symbols-outlined dd-icon">format_align_justify</span><span class="dd-text">Justify</span><span class="dd-hint">Ctrl+Shift+J</span></div>
          </div>
        </div>

        <span class="sep"></span>

        <button class="fb" (click)="exec('insertUnorderedList')" title="Bulleted list"><span class="material-symbols-outlined">format_list_bulleted</span></button>
        <button class="fb" (click)="exec('insertOrderedList')" title="Numbered list"><span class="material-symbols-outlined">format_list_numbered</span></button>
        
        <div class="menu-item" (click)="toggleMenu('indent', $event)" [class.active]="activeMenu === 'indent'" title="Indent">
          <span class="material-symbols-outlined" style="font-size: 20px;">format_indent_increase</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'indent'" style="min-width: 180px;">
            <div class="dd-item" (click)="exec('indent')"><span class="material-symbols-outlined dd-icon">format_indent_increase</span><span class="dd-text">Increase Indent</span><span class="dd-hint">Ctrl+M</span></div>
            <div class="dd-item" (click)="exec('outdent')"><span class="material-symbols-outlined dd-icon">format_indent_decrease</span><span class="dd-text">Decrease Indent</span><span class="dd-hint">Ctrl+Shift+M</span></div>
          </div>
        </div>

        <span class="sep"></span>

        <button class="fb" (mousedown)="$event.preventDefault()" (click)="insertLink()" title="Insert link"><span class="material-symbols-outlined">link</span></button>
        <div class="menu-item" (click)="toggleMenu('image', $event)" [class.active]="activeMenu === 'image'" title="Insert image">
          <span class="material-symbols-outlined" style="font-size: 20px;">image</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'image'" style="min-width: 200px;">
            <div class="dd-item" (click)="imageInput.click()"><span class="material-symbols-outlined dd-icon">upload</span><span class="dd-text">Upload from computer</span></div>
            <div class="dd-item" (click)="openImageModal('url')"><span class="material-symbols-outlined dd-icon">link</span><span class="dd-text">Insert a URL</span></div>
            <div class="dd-item" (click)="showToast('Google Photos')"><span class="material-symbols-outlined dd-icon">photo_library</span><span class="dd-text">Google Photos</span></div>
          </div>
        </div>
        <button class="fb" (click)="insertTable()" title="Insert table"><span class="material-symbols-outlined">table</span></button>
        <span class="sep"></span>
        <button class="fb" (click)="exec('superscript')" title="Superscript"><span style="font-family: serif; font-size: 14px; font-weight: 500;">A<sup>2</sup></span></button>
        <button class="fb" (click)="exec('subscript')" title="Subscript"><span style="font-family: serif; font-size: 14px; font-weight: 500;">A<sub>2</sub></span></button>
        <span class="sep"></span>
        <div class="menu-item" (click)="toggleMenu('case', $event)" [class.active]="activeMenu === 'case'" title="Change Case">
          <span style="font-family: serif; font-weight: 500; font-size: 15px; padding: 0 4px;">Ag</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'case'" style="min-width: 180px;">
            <div class="dd-item" (click)="changeCase('uppercase')"><span class="dd-text" style="text-transform: uppercase;">AG UPPERCASE</span><span class="dd-hint">Ctrl+Shift+A</span></div>
            <div class="dd-item" (click)="changeCase('lowercase')"><span class="dd-text" style="text-transform: lowercase;">ag lowercase</span></div>
            <div class="dd-item" (click)="changeCase('capitalize')"><span class="dd-text" style="text-transform: capitalize;">Ag Capitalize Each Word</span></div>
            <div class="dd-item" (click)="changeCase('sentence')"><span class="dd-text">Ag Sentence case</span></div>
            <div class="dd-item" (click)="changeCase('smallcaps')"><span class="dd-text" style="font-variant: small-caps;">Ag Small Caps</span></div>
          </div>
        </div>

        <div class="menu-item" (click)="toggleMenu('linespacing', $event)" [class.active]="activeMenu === 'linespacing'" title="Line & Paragraph Spacing">
          <span class="material-symbols-outlined" style="font-size: 20px;">format_line_spacing</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'linespacing'" style="min-width: 230px;">
            <div class="dd-item" (click)="setLineSpacing('1')"><span class="dd-text" style="padding-left: 28px;">1.0 (Single)</span></div>
            <div class="dd-item" (click)="setLineSpacing('1.2')"><span class="dd-text" style="padding-left: 28px;">1.2 (Normal)</span></div>
            <div class="dd-item" (click)="setLineSpacing('1.5')"><span class="dd-text" style="padding-left: 28px;">1.5</span></div>
            <div class="dd-item" (click)="setLineSpacing('2')"><span class="dd-text" style="padding-left: 28px;">2.0 (Double)</span></div>
            <div class="dd-sep"></div>
            <div class="dd-item" style="pointer-events: none; opacity: 0.7;"><span class="material-symbols-outlined dd-icon">format_letter_spacing</span><span class="dd-text">Character Spacing</span></div>
            <div class="dd-item" (click)="setCharSpacing('-1')"><span class="dd-text" style="padding-left: 28px;">Condensed</span></div>
            <div class="dd-item" (click)="setCharSpacing('0')"><span class="dd-text" style="padding-left: 28px;">Normal</span></div>
            <div class="dd-item" (click)="setCharSpacing('1')"><span class="dd-text" style="padding-left: 28px;">Expanded</span></div>
            <div class="dd-item" (click)="setCharSpacing('2')"><span class="dd-text" style="padding-left: 28px;">Wide</span></div>
          </div>
        </div>
        <span class="sep"></span>
        <button class="fb" (click)="toggleNonPrinting()" [class.active-fb]="showNonPrinting" title="Show non-printing characters"><span class="material-symbols-outlined">format_paragraph</span></button>
        <div class="menu-item" (click)="toggleMenu('direction', $event)" [class.active]="activeMenu === 'direction'" title="Text Direction">
          <span class="material-symbols-outlined" style="font-size: 20px;">format_textdirection_l_to_r</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'direction'" style="min-width: 160px; right: 0; left: auto;">
            <div class="dd-item" (click)="execDir('ltr')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ activeDir === 'ltr' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">format_textdirection_l_to_r</span><span class="dd-text">Left To Right</span></div>
            <div class="dd-item" (click)="execDir('rtl')"><span class="material-symbols-outlined dd-icon" style="font-size: 16px;">{{ activeDir === 'rtl' ? 'check' : '' }}</span><span class="material-symbols-outlined dd-icon">format_textdirection_r_to_l</span><span class="dd-text">Right To Left</span></div>
          </div>
        </div>
        <span class="sep"></span>
        <div class="menu-item" (click)="toggleMenu('advanced', $event)" [class.active]="activeMenu === 'advanced'" title="Advanced Layout">
          <span class="material-symbols-outlined" style="font-size: 20px;">format_shapes</span> <span class="material-symbols-outlined arrow-icon">expand_more</span>
          <div class="dropdown" *ngIf="activeMenu === 'advanced'" style="width: 320px; padding: 12px; right: 0; left: auto; cursor: default;" (click)="$event.stopPropagation()">
            <div style="font-size:11px; font-weight:600; color:#5f6368; text-transform:uppercase; margin-bottom:8px;">Drop Cap</div>
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px;">
              <div class="grid-tile" (click)="applyDropCap('none'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; color: #1a73e8; font-size:12px;">NONE</div>
              <div class="grid-tile" (click)="applyDropCap('style1'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; font-size: 24px; font-weight: bold;">D</div>
              <div class="grid-tile" (click)="applyDropCap('style2'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; font-size: 24px; font-weight: bold; background: #8ab4f8; color: white;">D</div>
              <div class="grid-tile" (click)="applyDropCap('style3'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; font-size: 24px; font-weight: bold; background: #81c995; color: white; border-radius: 12px;">D</div>
              <div class="grid-tile" (click)="applyDropCap('style4'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; font-size: 24px; font-weight: bold; border-bottom: 3px solid #8ab4f8;">D</div>
              <div class="grid-tile" (click)="applyDropCap('style5'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 12px; text-align: center; cursor: pointer; font-size: 24px; font-weight: bold; background: #f48fb1; color: white; border-top-right-radius: 24px; border-bottom-right-radius: 24px;">D</div>
            </div>
            <div style="font-size:11px; font-weight:600; color:#5f6368; text-transform:uppercase; margin: 16px 0 8px 0;">Block Quote</div>
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px;">
              <div class="grid-tile" (click)="applyBlockQuote('style1'); closeMenus()" style="border: 1px solid #dadce0; border-left: 4px solid #1a73e8; border-radius: 2px; padding: 8px; cursor: pointer; font-size: 9px; color: #5f6368; background: #e8f0fe;">Lorem ipsum dolor...</div>
              <div class="grid-tile" (click)="applyBlockQuote('style2'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 8px; cursor: pointer; font-size: 9px; color: #5f6368; background: #fce8e6;">Lorem ipsum dolor...</div>
              <div class="grid-tile" (click)="applyBlockQuote('style3'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 8px; cursor: pointer; font-size: 9px; color: #5f6368; background: #fef7e0;">Lorem ipsum dolor...</div>
              <div class="grid-tile" (click)="applyBlockQuote('style4'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 8px; cursor: pointer; font-size: 9px; color: #5f6368; background: #f3f3f3;">Lorem ipsum dolor...</div>
              <div class="grid-tile" (click)="applyBlockQuote('style5'); closeMenus()" style="border: 1px solid #dadce0; border-radius: 4px; padding: 8px; cursor: pointer; font-size: 9px; color: #ffffff; background: #3c4043;">Lorem ipsum dolor...</div>
              <div class="grid-tile" (click)="applyBlockQuote('style6'); closeMenus()" style="border: 1px solid #dadce0; border-left: 4px solid #00bcd4; border-radius: 4px; padding: 8px; cursor: pointer; font-size: 9px; color: #5f6368; background: #e0f7fa;">Lorem ipsum dolor...</div>
            </div>
          </div>
        </div>
      </div>

      <div class="equation-bar" *ngIf="showEquationToolbar">
         <button class="fb" (click)="insertText('+ï¿½')">+ï¿½</button> <button class="fb" (click)="insertText('+ï¿½')">+ï¿½</button> <button class="fb" (click)="insertText('+ï¿½')">+ï¿½</button> <span class="sep"></span>
         <button class="fb" (click)="insertText('Gï¿½ï¿½')">Gï¿½ï¿½</button> <button class="fb" (click)="insertText('Gï¿½')">Gï¿½</button> <button class="fb" (click)="insertText('Gï¿½ï¿½')">Gï¿½ï¿½</button> <span class="sep"></span>
         <button class="fb" (click)="insertText('Gï¿½P')">Gï¿½P</button> <button class="fb" (click)="insertText('Gï¿½ï¿½')">Gï¿½ï¿½</button> <button class="fb" (click)="insertText('Gï¿½ï¿½')">Gï¿½ï¿½</button>
      </div>

      <div class="editor-layout">
        <div class="sidebar" *ngIf="showPrintLayout">
          <div class="thumbnail" *ngFor="let p of pageCountArray; let i = index" (click)="scrollToPage(i)">
            <div class="thumb-page" style="overflow: hidden; position: relative; background: #fff; display: flex; flex-direction: column; align-items: center; padding-top: 10px; gap: 4px;">
              <!-- CSS Skeleton to look like a document preview without crashing the DOM -->
              <div style="width: 70%; height: 2px; background: #e0e0e0; border-radius: 1px;"></div>
              <div style="width: 80%; height: 2px; background: #e0e0e0; border-radius: 1px;"></div>
              <div style="width: 75%; height: 2px; background: #e0e0e0; border-radius: 1px;"></div>
              <div style="width: 85%; height: 2px; background: #e0e0e0; border-radius: 1px;"></div>
              <div style="width: 60%; height: 2px; background: #e0e0e0; border-radius: 1px;"></div>
              
              <button class="delete-page-btn" (click)="deletePage(i, $event)" title="Delete Page" style="z-index: 10;">&times;</button>
            </div>
            <div class="thumb-label">{{ i + 1 }}</div>
          </div>
          <div class="add-page-btn-wrapper" style="margin-top: 16px; display: flex; justify-content: center;">
            <button (click)="insertBreak()" title="Add new page" style="width: 40px; height: 40px; border-radius: 50%; background: #fff; border: 1px solid #dadce0; box-shadow: 0 1px 3px rgba(60,64,67,.1); display: flex; align-items: center; justify-content: center; cursor: pointer;">
              <span class="material-symbols-outlined" style="color: #1a73e8; font-size: 24px;">add</span>
            </button>
          </div>
        </div>
        <div class="page-area" #pageArea [class.no-print]="!showPrintLayout" (scroll)="updateOverlay()" (mousedown)="onPageMouseDown($event)">
          <div [style.width]="documentViewType === 'Web View' ? '100%' : '816px'" style="margin: 0 auto; display:flex; flex-direction: column; align-items: center; padding-bottom: 48px;" [style.zoom]="zoomLevel / 100">
            <div class="ruler" *ngIf="showRuler" [style.width]="documentViewType === 'Web View' ? '100%' : '816px'"></div>
            <div style="position: relative;" [style.width]="documentViewType === 'Web View' ? '100%' : '816px'">
              <div class="page-bg-layer" *ngIf="showPrintLayout">
                 <div class="page-bg" *ngFor="let p of pageCountArray"></div>
              </div>
              <div class="page" [attr.contenteditable]="viewMode === 'Editing' ? 'true' : 'false'" #editor
                [attr.spellcheck]="activeSpellErrors"
                [class.show-np]="showNonPrinting"
                [class.hide-images]="hideImages"
                [class.grid-lines]="showSmartGridLines"
                [class.object-indicator]="showObjectIndicator"
                [style.caret-color]="activeTrackChanges === 'On' ? activeMarkupColor : 'auto'"
                (contextmenu)="onContextMenu($event)"
                (paste)="onPaste($event)"
                (input)="onInput(editor)" (blur)="save()">
              </div>
            </div>
          </div>
        </div>

        <div class="sidebar right-sidebar" *ngIf="showAutomateSidebar" style="width: 300px; border-left: 1px solid #dadce0; background: #ffffff; padding: 24px 16px 20px 16px; overflow-y: auto; z-index: 10; position: relative; font-family: 'Inter', sans-serif;">
          <span class="material-symbols-outlined" style="position: absolute; top: 12px; right: 12px; cursor: pointer; font-size: 18px; color: #5f6368; padding: 4px; border-radius: 50%; transition: background 0.1s;" onmouseover="this.style.background='#f1f3f4'" onmouseout="this.style.background='transparent'" (click)="toggleAutomateSidebar('')">close</span>
          <div style="text-align: center; padding-bottom: 16px; margin-bottom: 20px; border-bottom: 1px solid #f1f3f4;">
            <div style="font-weight: 500; font-size: 16px; color: #202124;">{{ automateTitle }}</div>
            <div style="font-size: 12px; color: #5f6368; margin-top: 4px;">{{ automateSubtitle }}</div>
          </div>
          
          <div style="margin-bottom: 20px;">
            <div style="font-size: 13px; color: #3c4043; margin-bottom: 12px; line-height: 1.5;">{{ automateDescription }}</div>
            
            <button class="primary-btn" style="width: 100%; padding: 8px; border-radius: 4px; background: #1a73e8; color: #fff; border: none; font-weight: 500; cursor: pointer; margin-bottom: 12px;" (click)="showToast('Feature not yet connected to backend API')">Get Started</button>
            <button class="secondary-btn" style="width: 100%; padding: 8px; border-radius: 4px; background: #fff; color: #1a73e8; border: 1px solid #1a73e8; font-weight: 500; cursor: pointer;" (click)="toggleAutomateSidebar('')">Cancel</button>
          </div>
        </div>

        <div class="sidebar right-sidebar" *ngIf="showTextBoxOptions" style="width: 280px; border-left: 1px solid #dadce0; background: #ffffff; padding: 24px 16px 20px 16px; overflow-y: auto; z-index: 10; position: relative; font-family: 'Inter', sans-serif;">
          <span class="material-symbols-outlined" style="position: absolute; top: 12px; right: 12px; cursor: pointer; font-size: 18px; color: #5f6368; padding: 4px; border-radius: 50%; transition: background 0.1s;" onmouseover="this.style.background='#f1f3f4'" onmouseout="this.style.background='transparent'" (click)="toggleTextBoxSidebar(false)">close</span>
          <div style="text-align: center; padding-bottom: 16px; margin-bottom: 20px; border-bottom: 1px solid #f1f3f4;">
            <div style="font-weight: 500; font-size: 14px; color: #202124;">Text Box Options</div>
          </div>
          
          <div style="margin-bottom: 20px;">
            <div style="font-size: 12px; font-weight: 500; color: #5f6368; margin-bottom: 8px;">Alignment</div>
            <select [ngModel]="textBoxAlign" (ngModelChange)="textBoxAlign = $event" style="width: 100%; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; color: #202124; outline: none; background: #fff; cursor: pointer;">
              <option value="Top">Align Top</option>
              <option value="Middle">Align Middle</option>
              <option value="Bottom">Align Bottom</option>
            </select>
          </div>

          <div style="margin-bottom: 20px;">
            <div style="font-size: 12px; font-weight: 500; color: #5f6368; margin-bottom: 12px;">Style</div>
            <div style="display: flex; gap: 24px; align-items: center;">
              <div style="display: flex; gap: 8px; align-items: center;">
                <div style="font-size: 13px; color: #3c4043;">Border</div>
                <div class="menu-item" (click)="toggleMenu('boxBorderColor', $event)" [class.active]="activeMenu === 'boxBorderColor'" style="width: 28px; height: 28px; border: 1px solid #dadce0; border-radius: 4px; cursor: pointer; position: relative; display: flex; align-items: center; justify-content: center; padding: 0;">
                  <div [style.background]="textBoxBorderColor || 'transparent'" style="width: 18px; height: 18px; border-radius: 2px; border: 1px solid rgba(0,0,0,0.1);"></div>
                  <div class="dropdown" *ngIf="activeMenu === 'boxBorderColor'" style="min-width: 220px; padding:12px; cursor:default; right: 0; left: auto; top: 100%; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-radius: 8px;" (click)="$event.stopPropagation()">
                    <div class="dd-item" (click)="textBoxBorderColor = 'transparent'; closeMenus()" style="margin:0 -12px 12px -12px; border-bottom:1px solid #eee; padding:0 12px 12px 12px;"><span class="material-symbols-outlined dd-icon" style="color: #ea4335;">format_color_reset</span><span class="dd-text">No Border</span></div>
                    <div style="font-size:11px; font-weight:600; color:#5f6368; margin-bottom:8px;">Theme Colors</div>
                    <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="textBoxBorderColor = c; closeMenus()"></div></div>
                    <div class="cp-grid" style="margin-bottom: 12px;"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="textBoxBorderColor = c; closeMenus()"></div></div>
                    <div style="font-size:11px; font-weight:600; color:#5f6368; margin-bottom:8px;">Standard Colors</div>
                    <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="textBoxBorderColor = c; closeMenus()"></div></div>
                  </div>
                </div>
              </div>

              <div style="display: flex; gap: 8px; align-items: center;">
                <div style="font-size: 13px; color: #3c4043;">Fill</div>
                <div class="menu-item" (click)="toggleMenu('boxBgColor', $event)" [class.active]="activeMenu === 'boxBgColor'" style="width: 28px; height: 28px; border: 1px solid #dadce0; border-radius: 4px; cursor: pointer; position: relative; display: flex; align-items: center; justify-content: center; padding: 0;">
                  <div [style.background]="textBoxBgColor || 'transparent'" style="width: 18px; height: 18px; border-radius: 2px; border: 1px solid rgba(0,0,0,0.1);"></div>
                  <div class="dropdown" *ngIf="activeMenu === 'boxBgColor'" style="min-width: 220px; padding:12px; cursor:default; right: 0; left: auto; top: 100%; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-radius: 8px;" (click)="$event.stopPropagation()">
                    <div class="dd-item" (click)="textBoxBgColor = 'transparent'; closeMenus()" style="margin:0 -12px 12px -12px; border-bottom:1px solid #eee; padding:0 12px 12px 12px;"><span class="material-symbols-outlined dd-icon" style="color: #ea4335;">format_color_reset</span><span class="dd-text">No Fill</span></div>
                    <div style="font-size:11px; font-weight:600; color:#5f6368; margin-bottom:8px;">Theme Colors</div>
                    <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="textBoxBgColor = c; closeMenus()"></div></div>
                    <div class="cp-grid" style="margin-bottom: 12px;"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="textBoxBgColor = c; closeMenus()"></div></div>
                    <div style="font-size:11px; font-weight:600; color:#5f6368; margin-bottom:8px;">Standard Colors</div>
                    <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="textBoxBgColor = c; closeMenus()"></div></div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div style="margin-bottom: 24px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
              <div style="font-size: 12px; font-weight: 500; color: #5f6368;">Opacity</div>
              <div style="font-size: 12px; font-weight: 500; color: #3c4043;">{{ textBoxOpacity }}%</div>
            </div>
            <input type="range" min="0" max="100" [ngModel]="textBoxOpacity" (ngModelChange)="textBoxOpacity = $event" style="width: 100%; accent-color: #1a73e8; cursor: pointer; height: 4px;">
          </div>

          <div style="margin-bottom: 24px;">
            <div style="font-size: 12px; font-weight: 500; color: #5f6368; margin-bottom: 8px;">Dimensions</div>
            <div style="display: flex; gap: 12px;">
              <div style="flex: 1;">
                <div style="font-size: 11px; color: #80868b; margin-bottom: 4px;">W (px)</div>
                <input type="number" [ngModel]="textBoxWidth" (ngModelChange)="textBoxWidth = $event" style="width: 100%; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; outline: none;">
              </div>
              <div style="flex: 1;">
                <div style="font-size: 11px; color: #80868b; margin-bottom: 4px;">H (px)</div>
                <input type="number" [ngModel]="textBoxHeight" (ngModelChange)="textBoxHeight = $event" style="width: 100%; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; outline: none;">
              </div>
            </div>
          </div>
          
          <div style="margin-bottom: 12px;">
             <div style="font-size: 12px; font-weight: 500; color: #5f6368; margin-bottom: 8px;">Arrangement</div>
             <select [ngModel]="textBoxArrange" (ngModelChange)="textBoxArrange = $event" style="width: 100%; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; color: #202124; outline: none; background: #fff; cursor: pointer;">
               <option value="In front of text">Bring to Front (Over text)</option>
               <option value="Behind text">Send to Back (Behind text)</option>
               <option value="Inline">Inline with text</option>
             </select>
          </div>
        </div>
      </div>

      <input type="file" #imageInput (change)="onImageUpload($event)" accept="image/*" style="display: none">
      <input type="file" #videoInput (change)="onVideoUpload($event)" accept="video/*" style="display: none">
      <input type="file" #audioInput (change)="onAudioUpload($event)" accept="audio/*" style="display: none">
      <input type="file" #signatureInput (change)="onSignatureUpload($event)" accept="image/*" style="display: none">
      <input type="file" #docInput (change)="onDocOpen($event)" accept=".txt,.html" style="display: none">
      <input type="file" #importInput (change)="importFile($event)" style="display:none" accept=".docx,.pdf,.txt,.html,.rtf">
      <div class="toast" [class.show]="toastVisible">{{ toastMsg }}</div>

      <!-- Image Resize Overlay -->
      <div class="resize-overlay" *ngIf="selectedObject" 
           [style.top.px]="overlayRect.top" [style.left.px]="overlayRect.left" 
           [style.width.px]="overlayRect.width" [style.height.px]="overlayRect.height">
        <div class="resize-handle tl" (mousedown)="startResize($event, 'tl')"></div>
        <div class="resize-handle tr" (mousedown)="startResize($event, 'tr')"></div>
        <div class="resize-handle bl" (mousedown)="startResize($event, 'bl')"></div>
        <div class="resize-handle br" (mousedown)="startResize($event, 'br')"></div>
        
        <div class="image-actions" *ngIf="selectedObject.tagName === 'IMG'" style="position: absolute; top: -36px; right: -2px; display: flex; gap: 4px; background: white; border-radius: 4px; padding: 4px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); pointer-events: auto;">
           <button class="header-icon-btn" (click)="setWrapStyle('inline')" title="Inline"><span class="material-symbols-outlined" style="font-size: 18px;">format_align_justify</span></button>
           <button class="header-icon-btn" (click)="setWrapStyle('left')" title="Wrap Left"><span class="material-symbols-outlined" style="font-size: 18px;">format_align_left</span></button>
           <button class="header-icon-btn" (click)="setWrapStyle('right')" title="Wrap Right"><span class="material-symbols-outlined" style="font-size: 18px;">format_align_right</span></button>
           <button class="header-icon-btn" (click)="setWrapStyle('break')" title="Break Text"><span class="material-symbols-outlined" style="font-size: 18px;">wrap_text</span></button>
           <button class="header-icon-btn" (click)="setWrapStyle('absolute')" title="In Front of Text"><span class="material-symbols-outlined" style="font-size: 18px;">layers</span></button>
           <div style="width: 1px; background: #dadce0; margin: 2px 4px;"></div>
           <button class="header-icon-btn" (click)="openCropModal()" title="Crop Image"><span class="material-symbols-outlined" style="font-size: 18px;">crop</span></button>
        </div>
      </div>

      <!-- Crop Modal -->
      <div class="modal-overlay" *ngIf="cropModalVisible">
        <div class="modal" style="width: 540px; padding: 24px;">
          <h3 style="margin: 0 0 16px 0; font-size: 16px; text-align: left;">Crop Image</h3>
          <div style="position: relative; margin: 16px 0; background: #f1f3f4; border: 1px solid #dadce0; border-radius: 8px; display: flex; justify-content: center; align-items: center; min-height: 200px; padding: 20px;">
             <img #cropImageEl [src]="cropTargetSrc" style="max-width: 100%; max-height: 400px; display: block; user-select: none;" (load)="initCropArea()" draggable="false">
             <div class="crop-area" *ngIf="cropArea" 
                  [style.left.px]="cropArea.x" [style.top.px]="cropArea.y" 
                  [style.width.px]="cropArea.w" [style.height.px]="cropArea.h" 
                  style="position: absolute; border: 2px dashed #1a73e8; background: rgba(26,115,232,0.1); cursor: move; user-select: none;"
                  (mousedown)="startCropDrag($event)">
                <div class="resize-handle tl" (mousedown)="startCropResize($event, 'tl')" style="position: absolute; top: -7px; left: -7px; cursor: nwse-resize;"></div>
                <div class="resize-handle tr" (mousedown)="startCropResize($event, 'tr')" style="position: absolute; top: -7px; right: -7px; cursor: nesw-resize;"></div>
                <div class="resize-handle bl" (mousedown)="startCropResize($event, 'bl')" style="position: absolute; bottom: -7px; left: -7px; cursor: nesw-resize;"></div>
                <div class="resize-handle br" (mousedown)="startCropResize($event, 'br')" style="position: absolute; bottom: -7px; right: -7px; cursor: nwse-resize;"></div>
             </div>
          </div>
          <div class="modal-actions">
            <button class="btn outline" (click)="cropModalVisible = false">Cancel</button>
            <button class="btn" (click)="applyCrop()">Apply Crop</button>
          </div>
        </div>
      </div>

      <!-- Find Dialog -->
      <div class="find-dialog" *ngIf="findDialogVisible">
        <input type="text" [(ngModel)]="findTextQuery" placeholder="Find in document" (keyup.enter)="executeFind()">
        <button class="btn outline" (click)="executeFind()">Next</button>
        <button class="btn outline" (click)="findDialogVisible = false">Close</button>
      </div>

      <!-- Drawing Modal -->
      <div class="modal-overlay" *ngIf="drawingModalVisible">
        <div class="modal" style="width: 600px;">
          <h3>Drawing Canvas</h3>
          <canvas #drawCanvas width="560" height="300" style="border: 1px solid #ccc; cursor: crosshair; background: #fff;"
            (mousedown)="startDraw($event)" (mousemove)="draw($event)" (mouseup)="stopDraw()" (mouseleave)="stopDraw()"></canvas>
          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="clearDrawing()">Clear</button>
            <div style="flex: 1"></div>
            <button class="btn outline" (click)="drawingModalVisible = false">Cancel</button>
            <button class="btn" (click)="insertDrawingImage()">Save and Close</button>
          </div>
        </div>
      </div>

      <!-- Chart Modal -->
      <div class="modal-overlay" *ngIf="chartModalVisible">
        <div class="modal" style="width: 500px;">
          <h3>Chart Builder</h3>
          <p style="color: #5f6368; font-size: 13px;">Generate a dynamic chart from data</p>
          <div style="margin-top: 12px; display: flex; gap: 12px;">
            <div style="flex: 1; text-align: left;">
              <label style="display:block; margin-bottom: 4px; font-weight: 500;">Chart Type:</label>
              <select [(ngModel)]="chartConfig.type" class="form-control" style="width: 100%; padding: 8px; border: 1px solid #dadce0; border-radius: 4px;">
                <option value="Bar">Bar Chart</option>
                <option value="Line">Line Chart</option>
                <option value="Pie">Pie Chart</option>
                <option value="Donut">Donut Chart</option>
              </select>
            </div>
            <div style="flex: 1; text-align: left;">
               <label style="display:block; margin-bottom: 4px; font-weight: 500;">Base Color:</label>
               <input type="color" [(ngModel)]="chartConfig.color" style="width: 100%; height: 35px; border: 1px solid #dadce0; border-radius: 4px; cursor: pointer;">
            </div>
          </div>
          <div style="margin-top: 12px; text-align: left;">
            <label style="display:block; margin-bottom: 4px; font-weight: 500;">Chart Title:</label>
            <input type="text" [(ngModel)]="chartTitle" class="form-control" placeholder="e.g. Sales Report">
          </div>
          <div style="margin-top: 12px; text-align: left;">
            <label style="display:block; margin-bottom: 4px; font-weight: 500;">X-Axis Labels (comma separated):</label>
            <input type="text" [(ngModel)]="chartConfig.xAxisLabels" class="form-control" placeholder="e.g. Jan, Feb, Mar, Apr">
          </div>
          <div style="margin-top: 12px; text-align: left;">
            <label style="display:block; margin-bottom: 4px; font-weight: 500;">Y-Axis Data (comma separated numbers):</label>
            <input type="text" [(ngModel)]="chartValues" class="form-control" placeholder="e.g. 10, 25, 15, 30">
          </div>
          <div *ngIf="chartConfig.showZ" style="margin-top: 12px; text-align: left;">
            <label style="display:block; margin-bottom: 4px; font-weight: 500;">Z-Axis Data (comma separated numbers):</label>
            <input type="text" [(ngModel)]="chartConfig.zAxisValues" class="form-control" placeholder="e.g. 5, 10, 8, 12">
          </div>
          <div style="margin-top: 8px; text-align: left;">
            <button class="btn outline" style="padding: 4px 8px; font-size: 12px;" (click)="chartConfig.showZ = !chartConfig.showZ">
              {{ chartConfig.showZ ? '- Remove Z Axis' : '+ Add Z Axis' }}
            </button>
            <button class="btn outline" style="padding: 4px 8px; font-size: 12px; margin-left: 8px;" (click)="triggerExcelImport()">
              Import Excel Data
            </button>
          </div>
          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="chartModalVisible = false">Cancel</button>
            <button class="btn primary" (click)="generateAndInsertChart()">Insert Chart</button>
          </div>
        </div>
      </div>

      <!-- Generic Prompt Modal -->
      <div class="modal-overlay" *ngIf="promptModalVisible" style="z-index: 10000;">
        <div class="modal" style="width: 350px;">
          <h3 style="margin-top: 0; font-size: 16px; font-weight: 500;">{{ promptModalTitle }}</h3>
          <input type="text" [(ngModel)]="promptModalInput" style="width: 100%; padding: 8px; margin: 12px 0; border: 1px solid #dadce0; border-radius: 4px; font-family: 'Inter', sans-serif;" [placeholder]="promptModalPlaceholder">
          <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px;">
            <button class="btn" (click)="closePromptModal(false)">Cancel</button>
            <button class="btn primary" (click)="closePromptModal(true)">OK</button>
          </div>
        </div>
      </div>

      <!-- Generic Confirm Modal -->
      <div class="modal-overlay" *ngIf="confirmModalVisible" style="z-index: 10000;">
        <div class="modal" style="width: 350px;">
          <h3 style="margin-top: 0; font-size: 16px; font-weight: 500;">Confirm Action</h3>
          <div style="margin: 12px 0; color: #202124; font-size: 14px;">{{ confirmModalMessage }}</div>
          <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px;">
            <button class="btn" (click)="closeConfirmModal(false)">Cancel</button>
            <button class="btn primary" style="background: #ea4335;" (click)="closeConfirmModal(true)">Yes</button>
          </div>
        </div>
      </div>

      <!-- Image Connect Modal -->
      <div class="modal-overlay" *ngIf="imageModalVisible">
        <div class="modal">
          <h3 *ngIf="imageModalType === 'url' || imageModalType === 'web'">Insert Image from URL</h3>
          <h3 *ngIf="imageModalType === 'workdrive'">Connect to Zoho WorkDrive</h3>
          <h3 *ngIf="imageModalType === 'library'">Connect to My Library</h3>
          <h3 *ngIf="imageModalType === 'gphotos'">Connect to Google Photos</h3>
          <h3 *ngIf="imageModalType === 'flickr'">Connect to Flickr</h3>
          
          <div *ngIf="imageModalType === 'url' || imageModalType === 'web'" style="margin-top: 12px; text-align: left;">
            <label style="display:block; margin-bottom: 4px; font-weight: 500;">Image URL:</label>
            <input type="text" [(ngModel)]="imageUrlInput" class="form-control" placeholder="https://...">
          </div>
          
          <div *ngIf="imageModalType !== 'url' && imageModalType !== 'web'" style="margin-top: 12px; text-align: left; color: #5f6368; font-size: 13px;">
            To access your files from this service, you need to authenticate and authorize this application to view your media.
          </div>

          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="imageModalVisible = false">Cancel</button>
            <button class="btn primary" *ngIf="imageModalType === 'url' || imageModalType === 'web'" (click)="confirmImageUrl()">Insert</button>
            <button class="btn primary" *ngIf="imageModalType !== 'url' && imageModalType !== 'web'" (click)="showToast('Authentication pending...'); imageModalVisible = false;">Connect</button>
          </div>
        </div>
      </div>

      <!-- Cloud Import Modal -->
      <div class="modal-overlay" *ngIf="cloudModalOpen" (click)="cloudModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15); background: var(--bg-color, #fff);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 20px; font-size: 18px; color: var(--text-primary, #202124);">Connect Cloud Drive</h3>
          <ng-container *ngIf="!isConnectingCloud">
            <p style="font-size: 14px; color: var(--text-secondary, #5f6368); margin-bottom: 24px;">Select a cloud provider to authenticate and import documents into VMail Office Suite.</p>
            
            <div style="display: flex; flex-direction: column; gap: 12px;">
              <button class="btn" style="display: flex; align-items: center; justify-content: flex-start; gap: 16px; padding: 12px 16px; background: transparent; border: 1px solid var(--border-color, #dadce0); border-radius: 6px; color: var(--text-primary, #3c4043); font-weight: 500; cursor: pointer;" (click)="connectCloud('Google Drive')" onmouseover="this.style.background='var(--hover-bg, #f1f3f4)'" onmouseout="this.style.background='transparent'">
                <img src="https://upload.wikimedia.org/wikipedia/commons/d/da/Google_Drive_logo.png" style="width: 24px; height: 24px;"> Google Drive
              </button>
              <button class="btn" style="display: flex; align-items: center; justify-content: flex-start; gap: 16px; padding: 12px 16px; background: transparent; border: 1px solid var(--border-color, #dadce0); border-radius: 6px; color: var(--text-primary, #3c4043); font-weight: 500; cursor: pointer;" (click)="connectCloud('Dropbox')" onmouseover="this.style.background='var(--hover-bg, #f1f3f4)'" onmouseout="this.style.background='transparent'">
                <svg width="24" height="24" viewBox="0 0 43 40" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12.58 0L0 8.02l12.58 8.01 12.59-8.01L12.58 0zm17.84 0l-12.59 8.02 12.59 8.01L43 8.02 30.42 0zM0 24.05l12.58-8.02 12.58 8.02-12.58 8.01L0 24.05zm43 0l-12.58-8.02-12.59 8.02 12.59 8.01L43 24.05zM12.58 34.07l12.59 8.01 12.58-8.01-12.58-8.02-12.59 8.02z" fill="#0061FE"/></svg> Dropbox
              </button>
              <button class="btn" style="display: flex; align-items: center; justify-content: flex-start; gap: 16px; padding: 12px 16px; background: transparent; border: 1px solid var(--border-color, #dadce0); border-radius: 6px; color: var(--text-primary, #3c4043); font-weight: 500; cursor: pointer;" (click)="connectCloud('OneDrive')" onmouseover="this.style.background='var(--hover-bg, #f1f3f4)'" onmouseout="this.style.background='transparent'">
                <svg width="24" height="24" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg"><path d="M35 15C28 15 22 20 20 27 13 27 8 32 8 39 8 46 13 51 20 51L44 51C51 51 56 46 56 39 56 32 51 27 44 27 42 27 41 28 39 28 38 20 37 15 35 15Z" fill="#0078D4"/></svg> OneDrive
              </button>
            </div>
          </ng-container>
          
          <ng-container *ngIf="isConnectingCloud">
            <div style="padding: 40px 0; text-align: center; color: var(--text-secondary, #5f6368);">
              <div style="width: 24px; height: 24px; border: 3px solid #e8eaed; border-top: 3px solid #1a73e8; border-radius: 50%; margin: 0 auto 16px; animation: spin 1s linear infinite;"></div>
              <div style="font-weight: 500; color: var(--text-primary, #202124);">Connecting to {{connectingCloudName}}...</div>
              <div style="font-size: 12px; margin-top: 8px;">Waiting for authentication window</div>
            </div>
          </ng-container>
          
          <div style="margin-top: 24px; text-align: right;">
            <button class="btn" style="background: none; border: none; color: var(--text-secondary, #5f6368); padding: 8px 16px; cursor: pointer; font-weight: 600;" (click)="cloudModalOpen = false">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Share Modal -->
      <div class="modal-overlay" *ngIf="shareModalOpen" (click)="shareModalOpen = false">
        <div class="modal share-modal" (click)="$event.stopPropagation()">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:24px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <div style="background:#1a73e8; color:#fff; display:flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:4px;">
                <span class="material-symbols-outlined" style="font-size:16px;">description</span>
              </div>
              <h3 style="margin:0;">Share "{{ title || 'Untitled document' }}"</h3>
            </div>
            <button class="sm-close-btn" (click)="shareModalOpen = false" style="background:none; border:none; cursor:pointer;">
              <span class="material-symbols-outlined" style="font-size:20px;">close</span>
            </button>
          </div>
          <div style="position:relative; margin-bottom:32px;">
            <div style="display:flex; align-items:center; gap:12px; position:relative;">
              <div class="sm-input-box">
                <input type="text" class="sm-input" [(ngModel)]="shareQuery" (ngModelChange)="onShareSearch()" placeholder="Add people and groups">
                <div class="sm-dropdown-txt" (click)="shareRoleDropdownOpen = !shareRoleDropdownOpen" style="position:relative;">
                  {{ shareRole }} <span class="material-symbols-outlined" style="font-size:18px; color:inherit; opacity: 0.8;">arrow_drop_down</span>
                  
                  <div *ngIf="shareRoleDropdownOpen" class="sm-list" style="position:absolute; top:30px; right:0; left:auto; width:100px; z-index:100; min-width:100px; max-height:none;">
                     <div (click)="shareRole = 'View'; shareRoleDropdownOpen = false; $event.stopPropagation()" class="sm-list-item" style="padding:8px 12px; border-bottom:none;">View</div>
                     <div (click)="shareRole = 'Edit'; shareRoleDropdownOpen = false; $event.stopPropagation()" class="sm-list-item" style="padding:8px 12px; border-bottom:none;">Edit</div>
                  </div>
                </div>
              </div>
              <button (click)="performShare()" style="background:#1a73e8; color:#fff; border:none; border-radius:24px; font-weight:500; font-size:14px; padding:0 24px; height:44px; cursor:pointer; transition:background 0.2s;" onmouseover="this.style.background='#1557b0'" onmouseout="this.style.background='#1a73e8'">Share</button>
            </div>
            <div *ngIf="userSearchResults.length > 0" class="sm-list">
              <div *ngFor="let u of userSearchResults" (click)="selectShareUser(u)" class="sm-list-item">
                <div style="width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:14px; font-weight:500;" [style.background]="u.avatar_color">{{u.name.charAt(0).toUpperCase()}}</div>
                <div style="display:flex; flex-direction:column;">
                  <div class="name">{{u.name}}</div>
                  <div class="email">{{u.email}}</div>
                </div>
              </div>
            </div>
          </div>
          <div style="margin-bottom:32px;">
            <div style="font-size:11px; font-weight:600; color:#9aa0a6; letter-spacing:0.8px; margin-bottom:16px;">WHO CAN ACCESS</div>
            <div style="display:flex; align-items:center; justify-content:space-between;">
              <div style="display:flex; align-items:center; gap:16px;">
                <div class="sm-icon-bg">
                  <span class="material-symbols-outlined" style="font-size:20px;">{{ isPublic ? 'public' : 'link' }}</span>
                </div>
                <div>
                  <div class="sm-txt-main">{{ isPublic ? 'Public Link - Anyone on the internet can view' : 'Permalink - Private, not shared with anyone' }}</div>
                </div>
              </div>
              <button *ngIf="!isPublic" (click)="makePublic()" class="sm-sec-btn">
                <span class="material-symbols-outlined" style="font-size:16px;">settings</span> Make Public
              </button>
              <button *ngIf="isPublic" (click)="isPublic = false" class="sm-sec-btn">
                <span class="material-symbols-outlined" style="font-size:16px;">lock</span> Make Private
              </button>
            </div>
          </div>
          <div style="display:flex; align-items:center; justify-content:space-between;">
            <button (click)="copyLink()" class="sm-copy-btn">Copy Link</button>
            <button (click)="shareModalOpen = false" class="sm-done-btn">Done</button>
          </div>
        </div>
      </div>

      <!-- Details Modal -->
      <div class="modal-overlay" *ngIf="detailsModalOpen" (click)="detailsModalOpen = false">
        <div class="modal" style="width: 380px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 20px; font-size: 18px; color: #202124;">Document Details</h3>
          
          <div style="display: grid; grid-template-columns: 120px 1fr; gap: 12px 16px; font-size: 14px; color: #5f6368;">
            <strong style="color: #202124;">Title:</strong> <span>{{ title || 'Untitled document' }}</span>
            <strong style="color: #202124;">Location:</strong> <span>My Drive</span>
            <strong style="color: #202124;">Active Users:</strong> <span>{{ activeUsers }}</span>
            <strong style="color: #202124;">Words:</strong> <span>{{ detailsData.words }}</span>
            <strong style="color: #202124;">Characters:</strong> <span>{{ detailsData.chars }}</span>
          </div>

          <div style="margin-top: 24px; text-align: right;">
            <button class="btn" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;" (click)="detailsModalOpen = false">OK</button>
          </div>
        </div>
      </div>

      <!-- Trash Modal -->
      <div class="modal-overlay" *ngIf="trashModalOpen" (click)="trashModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Move to trash</h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">Are you sure you want to move "{{ title || 'Untitled document' }}" to the trash?</p>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="trashModalOpen = false" style="background: transparent; border: 1px solid #dadce0; color: #1a73e8; padding: 8px 24px; border-radius: 4px; font-weight: 500; cursor: pointer;">Cancel</button>
            <button class="btn" (click)="confirmTrashDoc()" style="background: #d93025; border: none; color: #fff; padding: 8px 24px; border-radius: 4px; font-weight: 500; cursor: pointer;">Move to trash</button>
          </div>
        </div>
      </div>

      <!-- Version History Modal -->
      <div class="modal-overlay" *ngIf="versionModalOpen" (click)="versionModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Version History</h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">Only the current active version is available for this document.<br>Auto-save is active.</p>
          <div style="display: flex; justify-content: flex-end;">
            <button class="btn" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;" (click)="versionModalOpen = false">OK</button>
          </div>
        </div>
      </div>

      <!-- Mark As Final Modal -->
      <div class="modal-overlay" *ngIf="markAsFinalModalOpen" (click)="markAsFinalModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Mark as Final</h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">This document will be marked as final to discourage editing. You can always edit it later.</p>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="markAsFinalModalOpen = false" style="background: transparent; border: 1px solid #dadce0; color: #1a73e8; padding: 8px 24px; border-radius: 4px;">Cancel</button>
            <button class="btn" (click)="markAsFinalModalOpen = false; showToast('Document marked as final')" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;">Mark Final</button>
          </div>
        </div>
      </div>

      <!-- Publish Modal -->
      <div class="modal-overlay" *ngIf="publishModalOpen" (click)="publishModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Publish to Web</h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">Make your content visible to anyone by publishing it to the web. You can get a link or embed code.</p>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="publishModalOpen = false" style="background: transparent; border: 1px solid #dadce0; color: #1a73e8; padding: 8px 24px; border-radius: 4px;">Cancel</button>
            <button class="btn" (click)="publishModalOpen = false; showToast('Document published')" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;">Publish</button>
          </div>
        </div>
      </div>

      <!-- Fill As Form Modal -->
      <div class="modal-overlay" *ngIf="fillFormModalOpen" (click)="fillFormModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Fill As Form</h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">No fillable fields detected in this document. Add form fields from the Automation menu first.</p>
          <div style="display: flex; justify-content: flex-end;">
            <button class="btn" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;" (click)="fillFormModalOpen = false">OK</button>
          </div>
        </div>
      </div>

      <!-- Page Setup Modal -->
      <div class="modal-overlay" *ngIf="pageSetupModalOpen" (click)="pageSetupModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Page Setup</h3>
          <div style="margin-bottom: 24px;">
             <label style="display:block; font-size:12px; color:#5f6368; margin-bottom:4px;">Orientation</label>
             <select style="width:100%; padding:8px; border:1px solid #dadce0; border-radius:4px; margin-bottom:16px;">
               <option>Portrait</option>
               <option>Landscape</option>
             </select>
             <label style="display:block; font-size:12px; color:#5f6368; margin-bottom:4px;">Paper Size</label>
             <select style="width:100%; padding:8px; border:1px solid #dadce0; border-radius:4px;">
               <option>A4 (21 cm × 29.7 cm)</option>
               <option>Letter (8.5" × 11")</option>
               <option>Legal (8.5" × 14")</option>
             </select>
          </div>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="pageSetupModalOpen = false" style="background: transparent; border: 1px solid #dadce0; color: #1a73e8; padding: 8px 24px; border-radius: 4px;">Cancel</button>
            <button class="btn" (click)="pageSetupModalOpen = false; showToast('Page settings saved')" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;">OK</button>
          </div>
        </div>
      </div>

      <!-- Save As Modal -->
      <div class="modal-overlay" *ngIf="saveAsModalOpen" (click)="saveAsModalOpen = false">
        <div class="modal" style="width: 380px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15); background: var(--bg-color, #fff);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 20px; font-size: 18px; color: var(--text-primary, #202124);">{{ saveAsType === 'template' ? 'Save As Template' : 'Save As' }}</h3>
          <div style="margin-bottom: 24px;">
             <label style="display: block; margin-bottom: 8px; font-size: 14px; font-weight: 500; color: var(--text-secondary, #3c4043);">Document Name</label>
             <input type="text" [(ngModel)]="saveAsInput" placeholder="Document name" style="width: 100%; padding: 10px; border: 1px solid var(--border-color, #dadce0); border-radius: 4px; box-sizing: border-box; outline: none; font-size: 14px; background: var(--bg-color, #fff); color: var(--text-primary, #202124);" (keyup.enter)="confirmSaveAs()" autofocus>
          </div>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="saveAsModalOpen = false" style="background: transparent; border: 1px solid var(--border-color, #dadce0); color: #1a73e8; padding: 8px 24px; border-radius: 4px;">Cancel</button>
            <button class="btn" (click)="confirmSaveAs()" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;">Save</button>
          </div>
        </div>
      </div>

      <!-- Password Modal -->
      <div class="modal-overlay" *ngIf="passwordModalOpen" (click)="passwordModalOpen = false">
        <div class="modal" style="width: 400px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 16px; font-size: 18px; color: #202124;">Password Protect Document</h3>
          <p style="margin: 0 0 16px 0; font-size: 14px; color: #5f6368;">Enter a password to encrypt this file. Anyone who tries to open it will need this password.</p>
          <div style="margin-bottom: 24px;">
             <input type="password" [(ngModel)]="passwordInput" placeholder="Enter password" style="width: 100%; padding: 10px; border: 1px solid #dadce0; border-radius: 4px; box-sizing: border-box; outline: none; font-size: 14px;" (keyup.enter)="submitPassword()" autofocus>
          </div>
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button class="btn" (click)="passwordModalOpen = false" style="background: transparent; border: 1px solid #dadce0; color: #1a73e8; padding: 8px 24px; border-radius: 4px;">Cancel</button>
            <button class="btn" (click)="submitPassword()" style="background: #1a73e8; color: #fff; border-radius: 4px; padding: 8px 24px;">Download</button>
          </div>
        </div>
      </div>
      <!-- Dictionary Modal -->
      <div class="modal-overlay" *ngIf="dictModalOpen">
        <div class="modal" style="width: 400px; max-height: 80vh; overflow-y: auto; text-align: left;">
          <h3 style="margin-top: 0;">Dictionary</h3>
          <div style="display: flex; gap: 8px; margin-top: 12px; margin-bottom: 16px;">
            <input type="text" [(ngModel)]="dictWord" class="form-control" placeholder="Enter a word" (keyup.enter)="searchDictionary()">
            <button class="btn" (click)="searchDictionary()">Search</button>
          </div>
          
          <div *ngIf="dictLoading" style="text-align: center; color: #5f6368; padding: 20px;">Searching...</div>
          <div *ngIf="dictError" style="color: #ea4335; text-align: center; padding: 20px;">{{ dictError }}</div>
          
          <div *ngIf="dictResults && dictResults.length > 0">
            <div *ngFor="let result of dictResults" style="margin-bottom: 16px;">
              <h4 style="margin: 0; font-size: 18px; color: #202124;">{{ result.word }} <span style="font-size: 14px; color: #5f6368; font-weight: normal;">{{ result.phonetic }}</span></h4>
              <div *ngFor="let meaning of result.meanings" style="margin-top: 8px;">
                <div style="font-style: italic; color: #1a73e8; font-size: 13px;">{{ meaning.partOfSpeech }}</div>
                <ul style="margin: 4px 0 0 0; padding-left: 20px; font-size: 13px; color: #3c4043;">
                  <li *ngFor="let def of meaning.definitions" style="margin-bottom: 4px;">
                    {{ def.definition }}
                    <div *ngIf="def.example" style="color: #5f6368; font-style: italic; margin-top: 2px;">"{{ def.example }}"</div>
                  </li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="dictModalOpen = false">Close</button>
          </div>
        </div>
      </div>
      
      <!-- Engagement Insights Modal -->
      <div class="modal-overlay" *ngIf="engagementModalOpen" (click)="engagementModalOpen = false">
        <div class="modal" style="width: 500px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 8px; font-size: 18px; color: #202124; display: flex; align-items: center; gap: 8px;">
            <span class="material-symbols-outlined" style="color: #1a73e8;">monitoring</span>
            Engagement Insights
          </h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">Real-time analytics for your document viewers.</p>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px;">
             <div style="padding: 16px; border: 1px solid #dadce0; border-radius: 8px; text-align: center;">
                <div style="font-size: 28px; font-weight: 600; color: #1a73e8; margin-bottom: 4px;">1,204</div>
                <div style="font-size: 12px; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px;">Total Views</div>
             </div>
             <div style="padding: 16px; border: 1px solid #dadce0; border-radius: 8px; text-align: center;">
                <div style="font-size: 28px; font-weight: 600; color: #34a853; margin-bottom: 4px;">4m 12s</div>
                <div style="font-size: 12px; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px;">Avg. Reading Time</div>
             </div>
             <div style="padding: 16px; border: 1px solid #dadce0; border-radius: 8px; text-align: center;">
                <div style="font-size: 28px; font-weight: 600; color: #fbbc04; margin-bottom: 4px;">89</div>
                <div style="font-size: 12px; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px;">Unique Visitors</div>
             </div>
             <div style="padding: 16px; border: 1px solid #dadce0; border-radius: 8px; text-align: center;">
                <div style="font-size: 28px; font-weight: 600; color: #ea4335; margin-bottom: 4px;">34%</div>
                <div style="font-size: 12px; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px;">Completion Rate</div>
             </div>
          </div>
          
          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="engagementModalOpen = false">Close Insights</button>
          </div>
        </div>

      </div>

      <!-- Word Count Modal -->
      <div class="modal-overlay" *ngIf="wordCountModalOpen" (click)="wordCountModalOpen = false">
        <div class="modal" style="width: 350px; padding: 24px; border-radius: 8px; text-align: left; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0; margin-bottom: 8px; font-size: 18px; color: #202124; display: flex; align-items: center; gap: 8px;">
            <span class="material-symbols-outlined" style="color: #1a73e8;">pin</span>
            Word count
          </h3>
          <p style="margin: 0 0 24px 0; font-size: 14px; color: #5f6368;">Current document statistics.</p>
          
          <div style="display: flex; flex-direction: column; gap: 12px; margin-bottom: 24px; border: 1px solid #dadce0; border-radius: 8px; overflow: hidden;">
             <div style="display: flex; justify-content: space-between; padding: 12px 16px; border-bottom: 1px solid #dadce0;">
                <span style="font-size: 14px; font-weight: 500; color: #3c4043;">Words</span>
                <span style="font-size: 14px; color: #5f6368;">{{ wcWords }}</span>
             </div>
             <div style="display: flex; justify-content: space-between; padding: 12px 16px;">
                <span style="font-size: 14px; font-weight: 500; color: #3c4043;">Characters</span>
                <span style="font-size: 14px; color: #5f6368;">{{ wcChars }}</span>
             </div>
          </div>
          
          <div class="modal-actions" style="margin-top: 16px;">
            <button class="btn outline" (click)="wordCountModalOpen = false">Close</button>
          </div>
        </div>
      </div>
      <!-- Custom Context Menu -->
      <div class="custom-context-menu" *ngIf="showContextMenu" [style.left.px]="contextMenuX" [style.top.px]="contextMenuY" (click)="$event.stopPropagation()">
         <ng-container *ngIf="isFetchingSuggestions">
            <div class="ccm-item" style="color: #666; font-style: italic;">Loading suggestions...</div>
            <div class="dd-sep"></div>
         </ng-container>
         <ng-container *ngIf="!isFetchingSuggestions && contextMenuSuggestions.length > 0">
            <div class="ccm-title">Suggestions</div>
            <div class="ccm-item suggestion-item" *ngFor="let sug of contextMenuSuggestions" (click)="applySuggestion(sug)">
               <span style="font-weight: 600; color: #1a73e8;">{{ sug }}</span>
            </div>
            <div class="dd-sep"></div>
         </ng-container>
         <ng-container *ngIf="!isFetchingSuggestions && contextMenuSuggestions.length === 0">
            <div class="ccm-item" style="color: #666; font-style: italic;">No suggestions</div>
            <div class="dd-sep"></div>
         </ng-container>
         <div class="ccm-item" (click)="execVal('cut', ''); showContextMenu = false"><span class="material-symbols-outlined dd-icon">content_cut</span> Cut</div>
         <div class="ccm-item" (click)="execVal('copy', ''); showContextMenu = false"><span class="material-symbols-outlined dd-icon">content_copy</span> Copy</div>
         <div class="ccm-item" (click)="execVal('paste', ''); showContextMenu = false"><span class="material-symbols-outlined dd-icon">content_paste</span> Paste</div>
      </div>

      <!-- Status Bar -->
      <div class="status-bar">
        <div class="status-left">
          <div class="status-item" (click)="toggleWidget('chat')">
            <span class="material-symbols-outlined" style="font-size:16px;color:#d32f2f;">chat</span>
            <span>Chats</span>
         </div>
          <div class="status-item" (click)="toggleWidget('channels')">
            <span class="material-symbols-outlined" style="font-size:16px;">group</span>
            <span>Channels</span>
         </div>
          <span class="status-divider"></span>
          <div class="status-info">Words: {{ wordCount }}</div>
          <div class="status-info">Chars: {{ charCount }}</div>
          <div class="status-info">Page: {{ currentPage }} of {{ pageCountArray.length }}</div>
        </div>
        <div class="status-right">
          <div class="status-info">{{ viewMode }}</div>
          <span class="status-divider"></span>
          <button class="zoom-btn" (click)="zoomOut()"><span class="material-symbols-outlined" style="font-size:16px;">remove</span></button>
          <div class="status-info">{{ zoomLevel }}%</div>
          <button class="zoom-btn" (click)="zoomIn()"><span class="material-symbols-outlined" style="font-size:16px;">add</span></button>
         </div>
      </div>

      <app-chat-widget [activeWidget]="activeWidget" (close)="activeWidget=null"></app-chat-widget>
      
      <!-- Loading Data Overlay -->
      <div class="loading-overlay" *ngIf="isLoadingDocument || isUploading">
        <div class="loading-modal shadow-lg" *ngIf="isLoadingDocument">
           <div class="lm-spinner"></div>
           <div class="lm-title">Loading Document...</div>
           <div class="lm-subtitle">Retrieving text, images, and formatting. Please wait.</div>
        </div>
        
        <div class="upload-modal shadow-lg" *ngIf="isUploading">
           <div class="um-icon">
              <span class="material-symbols-outlined" style="font-size:32px; color:#1a73e8;">cloud_upload</span>
           </div>
           <div class="um-title">Importing Document...</div>
           <div class="um-subtitle">Please wait while your file is securely uploaded and processed.</div>
           
           <div class="um-progress-container">
              <div class="um-progress-bar" [style.width]="uploadProgress + '%'"></div>
           </div>
           
           <div class="um-stats">
              <div class="um-percent">{{ uploadProgress }}%</div>
              <div class="um-time">{{ uploadTimeLeft }}</div>
           </div>
        </div>
      </div>
  `,
  styleUrls: ['./doc-editor.component.css']
})
export class DocEditorComponent implements OnInit, OnDestroy {
  contextMenuX = 0;
  contextMenuY = 0;
  
  isLoadingDocument = true;
  isUploading = false;
  uploadProgress = 0;
  uploadTimeLeft = '';
  private uploadStartTime = 0;
  
  sanitizeImportedHtml(html: string): string {
    if (!html) return html;
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    
    // Extract imgs from p tags because p tags cannot be split by autoPaginate
    const ps = Array.from(tempDiv.querySelectorAll('p'));
    for (const p of ps) {
       const imgs = Array.from(p.querySelectorAll('img'));
       if (imgs.length > 0) {
           for (const img of imgs) {
              const wrapper = document.createElement('div');
              wrapper.style.textAlign = 'center';
              img.style.maxWidth = '100%';
              wrapper.appendChild(img.cloneNode(true));
              p.parentNode?.insertBefore(wrapper, p);
              img.remove();
           }
           if (p.innerHTML.trim() === '') {
              p.remove();
           }
       }
    }
    return tempDiv.innerHTML;
  }

  showContextMenu = false;
  contextMenuSuggestions: string[] = [];
  isFetchingSuggestions = false;
  contextMenuRange: Range | null = null;
  private inputTimeout: any;
  private typingTimer: any;
  private hasReceivedInitialSync = false;
  @ViewChild('importInput') importInput!: ElementRef;
  @ViewChild('editor') editor!: ElementRef;
  recentDocs: any[] = [];
  activeWidget: string | null = null;
  toggleWidget(w: string) {
    if (this.activeWidget === w) this.activeWidget = null;
    else this.activeWidget = w;
  }
  docId = '';
  activeFontFamily = 'Arial';
  activeFontSize = 11;

  promptModalVisible = false;
  promptModalTitle = '';
  promptModalInput = '';
  promptModalPlaceholder = '';
  private promptModalCallback: ((value: string | null) => void) | null = null;

  showPrompt(title: string, defaultVal: string = '', placeholder: string = ''): Promise<string | null> {
    this.promptModalTitle = title;
    this.promptModalInput = defaultVal;
    this.promptModalPlaceholder = placeholder;
    this.promptModalVisible = true;
    return new Promise(resolve => {
      this.promptModalCallback = resolve;
    });
  }

  closePromptModal(confirm: boolean) {
    this.promptModalVisible = false;
    if (this.promptModalCallback) {
      this.promptModalCallback(confirm ? this.promptModalInput : null);
      this.promptModalCallback = null;
    }
  }

  confirmModalVisible = false;
  confirmModalMessage = '';
  private confirmModalCallback: ((value: boolean) => void) | null = null;

  showConfirm(message: string): Promise<boolean> {
    this.confirmModalMessage = message;
    this.confirmModalVisible = true;
    return new Promise(resolve => {
      this.confirmModalCallback = resolve;
    });
  }

  closeConfirmModal(confirm: boolean) {
    this.confirmModalVisible = false;
    if (this.confirmModalCallback) {
      this.confirmModalCallback(confirm);
      this.confirmModalCallback = null;
    }
  }

  async setPageBorders() {
    this.closeMenus();
    const border = await this.showPrompt('Enter border style (e.g. 2px solid #1a73e8):', '1px solid #000');
    if (border) {
      const pageEl = document.querySelector('.page') as HTMLElement;
      if (pageEl) {
        pageEl.style.border = border;
      }
    }
  }

  async setPageBackground() {
    this.closeMenus();
    const color = await this.showPrompt('Enter background color (e.g. #f1f3f4 or lightblue):', '#ffffff');
    if (color) {
      const pageEl = document.querySelector('.page') as HTMLElement;
      if (pageEl) {
        pageEl.style.backgroundColor = color;
      }
    }
  }

  async importDesign() {
    this.closeMenus();
    const url = await this.showPrompt('Enter URL of the design CSS file (e.g. https://example.com/theme.css):');
    if (url) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = url;
      document.head.appendChild(link);
      this.showToast('Design imported successfully');
    }
  }

  applyFontSet(font: string) {
    this.closeMenus();
    this.activeFontSet = font;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      pageEl.style.fontFamily = font;
    }
  }

  applyColorSet(theme: string) {
    this.closeMenus();
    this.activeColorSet = theme;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      if (theme === 'bold') {
        pageEl.style.color = '#333';
      } else if (theme === 'Treasury') {
        pageEl.style.color = '#3e1f1f';
      } else if (theme === 'Brushstrokes') {
        pageEl.style.color = '#5a4325';
      } else if (theme === 'Pinstripes') {
        pageEl.style.color = '#5a4342';
      } else {
        pageEl.style.color = '#000';
      }
    }
  }

  setPageSize(size: string) {
    this.closeMenus();
    this.activePageSize = size;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      if (size === 'A4') {
        pageEl.style.width = '210mm';
        pageEl.style.minHeight = '297mm';
      } else if (size === 'Letter') {
        pageEl.style.width = '8.5in';
        pageEl.style.minHeight = '11in';
      } else if (size === 'Legal') {
        pageEl.style.width = '8.5in';
        pageEl.style.minHeight = '14in';
      }
    }
  }

  setPageOrientation(orientation: string) {
    this.closeMenus();
    this.activeOrientation = orientation;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
       const currentWidth = pageEl.style.width || '8.5in';
       const currentHeight = pageEl.style.minHeight || '11in';
       if (orientation === 'Landscape' && !currentWidth.includes('11in')) {
          pageEl.style.width = '11in';
          pageEl.style.minHeight = '8.5in';
       } else if (orientation === 'Portrait') {
          pageEl.style.width = '8.5in';
          pageEl.style.minHeight = '11in';
       }
    }
  }

  applyDesign(designName: string) {
    this.closeMenus();
    this.activeCurrentDesign = designName;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      if (designName === 'The Writer') {
        this.applyFontSet('Roboto');
        this.applyColorSet('Default');
      } else if (designName === 'Modern Report') {
        this.applyFontSet('Istok Web');
        this.applyColorSet('Treasury');
      } else if (designName === 'Newsletter') {
        this.applyFontSet('Quicksand');
        this.applyColorSet('Brushstrokes');
      } else if (designName === 'Academic Paper') {
        this.applyFontSet('PT Sans');
        this.applyColorSet('bold');
      } else if (designName === 'Business Letter') {
        this.applyFontSet('Work Sans');
        this.applyColorSet('Pinstripes');
      } else if (designName === 'Creative Story') {
        this.applyFontSet('Cabin');
        this.applyColorSet('Default');
      }
      this.showToast(`Applied ${designName} design`);
    }
  }

  setPageColumns(columns: string) {
    this.closeMenus();
    this.activeColumns = columns;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      if (columns === 'One') {
        pageEl.style.columnCount = '1';
      } else if (columns === 'Two') {
        pageEl.style.columnCount = '2';
      } else if (columns === 'Three') {
        pageEl.style.columnCount = '3';
      } else if (columns === 'Left' || columns === 'Right') {
        pageEl.style.columnCount = '2';
        pageEl.style.columnRule = '1px solid #dadce0';
      }
    }
  }

  setPageMargins(margins: string) {
    this.closeMenus();
    this.activeMargins = margins;
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (pageEl) {
      if (margins === 'Normal') {
        pageEl.style.padding = '1in';
      } else if (margins === 'Narrow') {
        pageEl.style.padding = '0.5in';
      } else if (margins === 'Moderate') {
        pageEl.style.padding = '1in 0.75in';
      } else if (margins === 'Wide') {
        pageEl.style.padding = '1in 2.28in';
      } else if (margins === 'None') {
        pageEl.style.padding = '0in';
      }
    }
  }

  setAdvancedSetup(setup: string) {
    this.closeMenus();
    this.activeAdvancedSetup = setup;
    this.showToast(`Switched to ${setup}`);
  }

  setTrackChanges(state: string) {
    this.closeMenus();
    this.activeTrackChanges = state;
    if (state === 'On') {
      const el = document.querySelector('.page') as HTMLElement;
      if (el) el.focus();
      document.execCommand('foreColor', false, this.activeMarkupColor);
      document.execCommand('underline', false, 'true');
    } else {
      const el = document.querySelector('.page') as HTMLElement;
      if (el) el.focus();
      document.execCommand('foreColor', false, '#000000');
      document.execCommand('underline', false, 'false');
    }
    this.showToast(`Track Changes ${state}`);
  }

  setMarkupColor(color: string) {
    this.closeMenus();
    this.activeMarkupColor = color;
    if (this.activeTrackChanges === 'On') {
      const el = document.querySelector('.page') as HTMLElement;
      if (el) el.focus();
      document.execCommand('foreColor', false, color);
    }
    this.showToast('Markup Color updated');
  }

  toggleMarkupDefault() {
    this.markupDefault = !this.markupDefault;
  }

  toggleSpellSetting(setting: string) {
    this.closeMenus();
    if (setting === 'Spelling Errors') this.activeSpellErrors = !this.activeSpellErrors;
    if (setting === 'Grammar') this.activeGrammar = !this.activeGrammar;
    if (setting === 'Writing Quality') this.activeWritingQuality = !this.activeWritingQuality;
  }

  async setTransliteration(lang: string) {
    this.closeMenus();
    this.activeTransliteration = lang;
    
    if (lang === 'None') return;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && sel.toString().trim() !== '') {
      const word = sel.toString().trim();
      const savedRange = sel.getRangeAt(0).cloneRange();
      
      const translateMap: any = {
        'Bengali': 'bn',
        'Gujarati': 'gu',
        'Hindi': 'hi',
        'Kannada': 'kn',
        'Malayalam': 'ml',
        'Marathi': 'mr',
        'Odia': 'or',
        'Punjabi': 'pa',
        'Spanish': 'es',
        'French': 'fr',
        'Tamil': 'ta',
        'Telugu': 'te'
      };
      const code = translateMap[lang];
      if (code) {
        try {
          const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${code}&dt=t&q=${encodeURIComponent(word)}`;
          const res = await fetch(url);
          const data = await res.json();
          if (data && data[0]) {
             const translatedText = data[0].map((item: any) => item[0]).join('');
             savedRange.deleteContents();
             const textNode = document.createTextNode(translatedText);
             savedRange.insertNode(textNode);
             
             const newRange = document.createRange();
             newRange.setStartAfter(textNode);
             newRange.collapse(true);
             sel.removeAllRanges();
             sel.addRange(newRange);
          }
        } catch (err) {
          console.error('Translation of selected text failed:', err);
        }
      }
    }
  }

  async insertCode() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    
    if (savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);

      const div = document.createElement('div');
      div.contentEditable = 'false';
      div.style.background = '#f1f3f4';
      div.style.border = '1px solid #dadce0';
      div.style.borderRadius = '6px';
      div.style.padding = '12px';
      div.style.margin = '16px 0';
      div.style.fontFamily = "'Courier New', Courier, monospace";
      div.style.fontSize = '13px';
      div.style.color = '#202124';
      div.style.position = 'relative';
      
      const header = document.createElement('div');
      header.style.fontSize = '11px';
      header.style.color = '#5f6368';
      header.style.marginBottom = '8px';
      header.style.textTransform = 'uppercase';
      header.style.fontFamily = "'Inter', sans-serif";
      header.innerText = 'Code Block';
      
      const code = document.createElement('div');
      code.contentEditable = 'true';
      code.style.outline = 'none';
      code.style.minHeight = '20px';
      code.style.whiteSpace = 'pre-wrap';
      code.innerText = '// Write your code here...';
      
      div.appendChild(header);
      div.appendChild(code);
      
      savedRange.insertNode(div);
      
      // Insert a br after to allow typing outside the block
      const br = document.createElement('br');
      if (div.parentNode) {
        div.parentNode.insertBefore(br, div.nextSibling);
      }
      
      // Move caret into the code block so they can start typing immediately
      const newRange = document.createRange();
      newRange.selectNodeContents(code);
      newRange.collapse(false);
      sel?.removeAllRanges();
      sel?.addRange(newRange);
      code.focus();
    }
  }

  async insertComment() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    if (!savedRange || savedRange.collapsed) {
      this.showToast('Please select some text to comment on.');
      return;
    }
    const comment = await this.showPrompt('Enter your comment:');
    if (comment && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const selectedText = savedRange.toString();
      const html = `<span class="doc-comment-span" data-comment="${comment.replace(/"/g, '&quot;')}" style="background-color: #fce8e6; border-bottom: 2px solid #ea4335; cursor: pointer;" title="Click to view comment">${selectedText}</span>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  async lockUnlockContent() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    if (savedRange && !savedRange.collapsed) {
      const div = document.createElement('div');
      div.appendChild(savedRange.cloneContents());
      
      if (div.querySelector('.locked-content')) {
        const page = document.querySelector('.page') as HTMLElement;
        const lockedElements = page.querySelectorAll('.locked-content');
        let unlockedCount = 0;
        lockedElements.forEach(node => {
          if (sel?.containsNode(node, true)) {
            const textNode = document.createTextNode(node.textContent || '');
            node.parentNode?.replaceChild(textNode, node);
            unlockedCount++;
          }
        });
        if (unlockedCount > 0) {
          this.showToast(`Unlocked ${unlockedCount} section(s).`);
          return;
        }
      }

      const container = savedRange.commonAncestorContainer;
      const el = container.nodeType === 3 ? container.parentElement : container as HTMLElement;
      const lockedNode = el?.closest('.locked-content') as HTMLElement;
      if (lockedNode) {
         const text = document.createTextNode(lockedNode.textContent || '');
         lockedNode.parentNode?.replaceChild(text, lockedNode);
         this.showToast('Content Unlocked.');
         return;
      }

      const html = `<span class="locked-content" contenteditable="false" style="background:#f1f3f4; border: 1px dashed #ccc;" title="Locked Content">${savedRange.toString()}</span>&#8203;`;
      document.execCommand('insertHTML', false, html);
    } else {
      this.showToast('Please select some text to lock/unlock.');
    }
  }

  async maskContent() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    if (savedRange && !savedRange.collapsed) {
      const div = document.createElement('div');
      div.appendChild(savedRange.cloneContents());
      
      if (div.querySelector('.masked-content')) {
        const page = document.querySelector('.page') as HTMLElement;
        const maskedElements = page.querySelectorAll('.masked-content');
        let unmaskedCount = 0;
        maskedElements.forEach(node => {
          if (sel?.containsNode(node, true)) {
            const textNode = document.createTextNode(node.textContent || '');
            node.parentNode?.replaceChild(textNode, node);
            unmaskedCount++;
          }
        });
        if (unmaskedCount > 0) {
          this.showToast(`Unmasked ${unmaskedCount} section(s).`);
          return;
        }
      }

      const container = savedRange.commonAncestorContainer;
      const el = container.nodeType === 3 ? container.parentElement : container as HTMLElement;
      const maskedNode = el?.closest('.masked-content') as HTMLElement;
      if (maskedNode) {
         const text = document.createTextNode(maskedNode.textContent || '');
         maskedNode.parentNode?.replaceChild(text, maskedNode);
         this.showToast('Content Unmasked.');
         return;
      }

      const html = `<span class="masked-content" contenteditable="false" style="background-color: black; color: black;" title="Masked Content">${savedRange.toString()}</span>&#8203;`;
      document.execCommand('insertHTML', false, html);
    } else {
      this.showToast('Please select some text to mask/unmask.');
    }
  }

  async compareVersions() {
    this.closeMenus();
    const ver = await this.showPrompt('Enter Document Version/ID to compare:', 'v1.0');
    if (ver) {
      this.showToast(`Comparing current document with version ${ver}...`);
    }
  }

  async combineRevisions() {
    this.closeMenus();
    const rev = await this.showPrompt('Enter Document Version/ID to combine:', 'v1.0');
    if (rev) {
      this.showToast(`Combining revisions from version ${rev}...`);
    }
  }

  toggleNotificationSettings() {
    this.closeMenus();
    this.activeNotifications = !this.activeNotifications;
    this.showToast(this.activeNotifications ? 'Notifications Enabled' : 'Notifications Disabled');
  }

  async insertVideo() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const url = await this.showPrompt('Enter Video URL (YouTube, Vimeo, etc):', 'https://');
    if (url && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      let embedUrl = url;
      if (url.includes('youtube.com/watch?v=')) {
        embedUrl = url.replace('watch?v=', 'embed/');
      } else if (url.includes('youtu.be/')) {
        embedUrl = url.replace('youtu.be/', 'youtube.com/embed/');
      }
      const html = `<div contenteditable="false" style="display: inline-block; margin: 10px 0;"><iframe width="560" height="315" src="${embedUrl}" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen style="max-width: 100%;"></iframe></div><br>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  async insertAudio() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const url = await this.showPrompt('Enter Audio File URL (.mp3, .wav):', 'https://');
    if (url && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const html = `<div contenteditable="false" style="display: inline-block; margin: 10px 0;"><audio controls src="${url}" style="width: 100%; max-width: 300px;"></audio></div><br>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  async insertQRCode() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const data = await this.showPrompt('Enter text or URL to encode as QR Code:');
    if (data && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const html = `<img src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(data)}" alt="QR Code" style="vertical-align: middle; margin: 8px;" /><br>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  async insertBarcode() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const data = await this.showPrompt('Enter data to encode as Barcode:');
    if (data && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const html = `<img src="https://barcode.tec-it.com/barcode.ashx?data=${encodeURIComponent(data)}&code=Code128&translate-esc=true" alt="Barcode" style="vertical-align: middle; margin: 8px;" /><br>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  insertHeader() {
    this.closeMenus();
    const html = `<div style="border-bottom: 1px solid #dadce0; padding-bottom: 8px; margin-bottom: 16px; color: #5f6368; font-size: 11px;">[Header Content]</div>`;
    document.execCommand('insertHTML', false, html);
  }

  insertFooter() {
    this.closeMenus();
    const html = `<div style="border-top: 1px solid #dadce0; padding-top: 8px; margin-top: 16px; color: #5f6368; font-size: 11px;">[Footer Content]</div>`;
    document.execCommand('insertHTML', false, html);
  }

  insertPageNumbers() {
    this.closeMenus();
    const html = `<div style="border-top: 1px solid #dadce0; padding-top: 8px; margin-top: 16px; color: #5f6368; font-size: 11px; text-align: right;">Page <span class="page-number-placeholder">1</span></div>`;
    document.execCommand('insertHTML', false, html);
  }
  activeBlockStyle = 'Normal text';
  fonts = ['Arial', 'Caveat', 'Comfortaa', 'Comic Sans MS', 'Courier New', 'EB Garamond', 'Georgia', 'Impact', 'Lexend', 'Lobster', 'Lora', 'Merriweather', 'Oswald', 'Pacifico', 'Playfair Display', 'Roboto', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

  title = 'Untitled';
  htmlContent = '';
  activeUsers = 1;
  activeMenu: string | null = null;
  activeFontSet: string = 'Roboto';
  activeColorSet: string = 'Default';
  activePageSize: string = 'Letter';
  activeOrientation: string = 'Portrait';
  activeColumns: string = 'One';
  activeMargins: string = 'Normal';
  activeAdvancedSetup: string = 'Document-Level Setup';
  activeCurrentDesign: string = 'The Writer';
  
  activeTrackChanges: string = 'Off';
  activeMarkupColor: string = '#34a853';
  markupColors: string[] = ['#34a853', '#9c27b0', '#a1887f', '#26a69a', '#689f38', '#039be5', '#d32f2f', '#003cff', '#fbc02d', '#bcaaa4', '#81d4fa', '#1a237e', '#8d6e63', '#e91e63', '#f48fb1', '#ce93d8', '#ff0000', '#ff6d00'];
  markupDefault: boolean = false;
  
  activeSpellErrors: boolean = true;
  activeGrammar: boolean = true;
  activeWritingQuality: boolean = true;
  activeLanguage: string = 'English (US)';
  
  activeTransliteration: string = 'None';
  activeNotifications: boolean = true;
  
  wordCount = 0;
  charCount = 0;
  currentPage = 1;
  zoomLevel = 100;
  isSaving = false;
  autoSaveEnabled = true;

  tableGridRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  tableGridCols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  hoveredTableRow = 0;
  hoveredTableCol = 0;

  insertGridTable(rows: number, cols: number) {
    this.closeMenus();
    let html = '<table border="1" style="border-collapse: collapse; width: 100%; margin: 10px 0;"><tbody>';
    for (let r = 0; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        html += '<td style="padding: 4px;"><br></td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table>';

    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    document.execCommand('insertHTML', false, html);
  }

  insertTextBox() {
    this.closeMenus();
    const html = `<div class="vmail-text-box" style="box-sizing: border-box; border:1px solid #000; padding:8px; display:inline-block; width: 150px; height: 50px; overflow:hidden;" contenteditable="true">Text Box</div>&nbsp;`;
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    document.execCommand('insertHTML', false, html);
    this.showTextBoxOptions = true;
  }

  insertQuickPart() {
    this.closeMenus();
    const html = `<div style="margin-top: 20px;"><p>Sincerely,</p><br><p><strong>${this.auth.user?.name || 'User'}</strong></p></div><br>`;
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    document.execCommand('insertHTML', false, html);
  }

  async saveQuickPart() {
    this.closeMenus();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) {
      this.showToast('Please select some text or elements first.');
      return;
    }
    const html = sel.getRangeAt(0).cloneContents();
    const div = document.createElement('div');
    div.appendChild(html);
    const content = div.innerHTML;
    
    const name = await this.showPrompt('Enter a name for this Quick Part:');
    if (name) {
      this.showToast(`Saved selection as Quick Part: ${name}`);
    }
  }
  lastSavedTime: Date | null = null;
  isStarred = false;

  themeColorsTop = [
    '#000000', '#434343', '#666666', '#999999', '#cccccc', '#efefef', '#f3f3f3', '#ffffff', '#ff0000', '#00ff00'
  ];
  themeColorsGrid = [
    '#f2f2f2', '#7f7f7f', '#d0cece', '#d6dce4', '#d9e1f2', '#fce4d6', '#ededed', '#fff2cc', '#deebf7', '#e2efda',
    '#d8d8d8', '#595959', '#a2a2a2', '#adb9ca', '#b4c6e7', '#f8cbad', '#dbdbdb', '#ffe699', '#bdd7ee', '#c6e0b4',
    '#bfbfbf', '#3f3f3f', '#7b7b7b', '#8497b0', '#8ea9db', '#f4b084', '#c9c9c9', '#ffd966', '#9dc3e6', '#a9d08e',
    '#a5a5a5', '#262626', '#525252', '#333f4f', '#2f5597', '#c55a11', '#7b7b7b', '#bf8f00', '#2e75b6', '#548235',
    '#7f7f7f', '#0c0c0c', '#252525', '#222a35', '#1f3864', '#833c0c', '#525252', '#7f6000', '#1e4e79', '#375623'
  ];
  standardColors = [
    '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'
  ];

  get savedTimeStr() {
    if (!this.lastSavedTime) return 'Saved';
    return 'Saved at ' + this.lastSavedTime.toLocaleTimeString();
  }

  isBold = false;
  isItalic = false;
  isUnderline = false;
  isStrikethrough = false;
  activeColor = '#000000';
  activeHighlight = '#ffffff';

  get currentUrl(): string { return window.location.href; }
  
  viewMode: 'Editing' | 'Viewing' = 'Editing';
  showPrintLayout = true;
  showTextBoxOptions = false;
  showAutomateSidebar = false;
  automateTitle = '';
  automateSubtitle = '';
  automateDescription = '';
  showRuler = false;
  showEquationToolbar = false;
  showNonPrinting = false;

  shareModalOpen = false;
  detailsModalOpen = false;
  trashModalOpen = false;
  versionModalOpen = false;
  publishModalOpen = false;
  markAsFinalModalOpen = false;
  pageSetupModalOpen = false;
  fillFormModalOpen = false;
  passwordModalOpen = false;
  passwordInput = '';
  passwordFormat = '';
  saveAsModalOpen = false;
  saveAsInput = '';
  saveAsType = 'document';
  cloudModalOpen = false;
  isConnectingCloud = false;
  connectingCloudName = '';
  documentViewType = 'Page View';
  showBookmarks = false;
  showSmartGridLines = false;
  showObjectIndicator = true;
  appearanceMode = 'System Default';
  hideImages = false;
  isReaderView = false;
  showNavigator = false;
  activeCharSpacing = 'Normal';
  activeDir = 'ltr';

  get isDarkMode() {
    if (this.appearanceMode === 'Dark Mode') return true;
    if (this.appearanceMode === 'System Default') {
      return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    }
    return false;
  }
  detailsData = { words: 0, chars: 0 };
  isPublic = false;
  shareQuery = '';
  shareRoleDropdownOpen = false;
  shareRole = 'View';
  userSearchResults: any[] = [];
  findDialogVisible = false;

  cropModalVisible = false;
  cropTargetSrc = '';
  cropTargetImg: HTMLImageElement | null = null;
  cropArea: { x: number, y: number, w: number, h: number } | null = null;
  cropAction: string | null = null;
  cropStartX = 0;
  cropStartY = 0;
  cropStartArea: any = null;
  @ViewChild('cropImageEl') cropImageEl!: ElementRef<HTMLImageElement>;
  findTextQuery = '';
  pageCountArray: number[] = [0];
  toastVisible = false;
  toastMsg = '';
  activeAutocorrect = true;

  translateSavedRange: Range | null = null;
  dictModalOpen = false;
  engagementModalOpen = false;
  wordCountModalOpen = false;
  wcWords = 0;
  wcChars = 0;
  dictWord = '';
  dictLoading = false;
  dictError = '';
  dictResults: any[] = [];
  private syncSub?: Subscription;
  private applyingRemote = false;

  selectedObject: HTMLElement | null = null;

  rgbToHex(rgb: string): string {
    if (!rgb || rgb === 'transparent') return '#ffffff';
    if (rgb.startsWith('#')) return rgb;
    const result = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(rgb);
    if (result) {
      return '#' + (1 << 24 | parseInt(result[1]) << 16 | parseInt(result[2]) << 8 | parseInt(result[3])).toString(16).slice(1);
    }
    return '#ffffff';
  }

  get textBoxAlign(): string {
    if (!this.selectedObject) return 'Top';
    const jc = this.selectedObject.style.justifyContent;
    return jc === 'center' ? 'Middle' : jc === 'flex-end' ? 'Bottom' : 'Top';
  }
  set textBoxAlign(val: string) {
    if (this.selectedObject) {
      this.selectedObject.style.display = 'inline-flex';
      this.selectedObject.style.flexDirection = 'column';
      this.selectedObject.style.justifyContent = val === 'Middle' ? 'center' : val === 'Bottom' ? 'flex-end' : 'flex-start';
      this.save();
    }
  }

  get textBoxOpacity(): number {
    if (!this.selectedObject) return 100;
    return Math.round(parseFloat(this.selectedObject.style.opacity || '1') * 100);
  }
  set textBoxOpacity(val: number) {
    if (this.selectedObject) {
      this.selectedObject.style.opacity = (val / 100).toString();
      this.save();
    }
  }

  get textBoxWidth(): number {
    if (!this.selectedObject) return 150;
    return this.selectedObject.offsetWidth || parseFloat(this.selectedObject.style.width || '150');
  }
  set textBoxWidth(val: number) {
    if (this.selectedObject) {
      this.selectedObject.style.width = val + 'px';
      this.updateOverlay();
      this.save();
    }
  }

  get textBoxHeight(): number {
    if (!this.selectedObject) return 50;
    return this.selectedObject.offsetHeight || parseFloat(this.selectedObject.style.height || '50');
  }
  set textBoxHeight(val: number) {
    if (this.selectedObject) {
      this.selectedObject.style.height = val + 'px';
      this.updateOverlay();
      this.save();
    }
  }

  get textBoxBgColor(): string {
    if (!this.selectedObject) return '#ffffff';
    return this.rgbToHex(this.selectedObject.style.backgroundColor) || '#ffffff';
  }
  set textBoxBgColor(val: string) {
    if (this.selectedObject) {
      this.selectedObject.style.backgroundColor = val;
      this.save();
    }
  }

  get textBoxBorderColor(): string {
    if (!this.selectedObject) return '#000000';
    return this.rgbToHex(this.selectedObject.style.borderColor) || '#000000';
  }
  set textBoxBorderColor(val: string) {
    if (this.selectedObject) {
      this.selectedObject.style.borderColor = val;
      this.save();
    }
  }

  get textBoxArrange(): string {
    if (!this.selectedObject) return 'Inline';
    if (this.selectedObject.style.position === 'absolute') {
      return this.selectedObject.style.zIndex === '-1' ? 'Behind text' : 'In front of text';
    }
    return 'Inline';
  }
  set textBoxArrange(val: string) {
    if (this.selectedObject) {
      if (val === 'Inline') {
        this.selectedObject.style.position = 'relative';
        this.selectedObject.style.zIndex = 'auto';
        this.selectedObject.style.left = 'auto';
        this.selectedObject.style.top = 'auto';
      } else {
        if (this.selectedObject.style.position !== 'absolute') {
          const pageEl = document.querySelector('.page') as HTMLElement;
          if (pageEl) {
            const pageRect = pageEl.getBoundingClientRect();
            const imgRect = this.selectedObject.getBoundingClientRect();
            this.selectedObject.style.position = 'absolute';
            this.selectedObject.style.left = (imgRect.left - pageRect.left) + 'px';
            this.selectedObject.style.top = (imgRect.top - pageRect.top) + 'px';
          } else {
            this.selectedObject.style.position = 'absolute';
          }
        }
        this.selectedObject.style.zIndex = val === 'Behind text' ? '-1' : '1';
      }
      this.updateOverlay();
      this.save();
    }
  }
  overlayRect = { top: 0, left: 0, width: 0, height: 0 };
  isResizing = false;
  resizeCorner = 'br';
  startX = 0;
  startY = 0;
  startW = 0;
  startH = 0;

  isDraggingImage = false;
  dragStartX = 0;
  dragStartY = 0;
  imgStartX = 0;
  imgStartY = 0;
  nativeDragStartWidth = 0;
  nativeDragStartHeight = 0;

  // Drawing Modal State
  drawingModalVisible = false;
  isDrawing = false;
  drawCtx: CanvasRenderingContext2D | null = null;
  lastX = 0;
  lastY = 0;
  @ViewChild('drawCanvas') drawCanvasRef!: ElementRef<HTMLCanvasElement>;

  // Chart Modal State
  chartModalVisible = false;
  chartTitle = 'Quarterly Revenue';
  chartValues = '40, 80, 60, 100';
  chartConfig = {
    type: 'Bar',
    xAxisLabels: 'Q1, Q2, Q3, Q4',
    zAxisValues: '',
    color: '#1a73e8',
    showZ: false
  };
  imageModalVisible = false;
  imageModalType = '';
  imageUrlInput = '';

  @ViewChild('pageArea') pageAreaRef!: ElementRef<HTMLElement>;

  get initials() {
    return (this.auth.user?.name ?? 'U').split(' ').map((w: string) => w[0]).join('').toUpperCase().slice(0, 2);
  }

  constructor(
    private route: ActivatedRoute,
    private router: Router,
    private api: ApiService,
    public auth: AuthService
  ) { }

  private globalClickHandler = (e: MouseEvent) => {
    const target = e.target as HTMLElement;
    if (target && target.classList.contains('doc-comment-span')) {
      const comment = target.getAttribute('data-comment');
      if (comment) {
        this.showToast(`Comment: ${comment}`);
      }
    }
  };
  
  goHome() {
    window.location.href = 'https://docs.vsnaptechnology.com/';
  }

  ngOnInit() {
    this.docId = this.route.snapshot.paramMap.get('id') ?? '';
    document.execCommand('defaultParagraphSeparator', false, 'div');
    document.addEventListener('click', this.globalClickHandler);
    
    this.api.listDocuments().subscribe(docs => {
      this.recentDocs = docs.filter((d: any) => d.doc_type === 'writer' && d.id !== this.docId).slice(0, 5);
    });

    this.api.getDocument(this.docId).subscribe((doc: any) => {
      this.title = doc.title;
      // ── Set browser tab title + favicon ──────────────────────────────
      document.title = (doc.title || 'Untitled') + ' - Vsnap Writer';
      const docFavicon = `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%234285F4'/><rect x='7' y='7' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='12' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='17' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='22' width='11' height='2.5' rx='1.2' fill='white'/></svg>`;
      let docLink: HTMLLinkElement = document.querySelector("link[rel~='icon']") as HTMLLinkElement;
      if (!docLink) { docLink = document.createElement('link'); docLink.rel = 'icon'; document.head.appendChild(docLink); }
      docLink.href = docFavicon;
      // ─────────────────────────────────────────────────────────────────
      if (doc.updated_at) {
        this.lastSavedTime = new Date(doc.updated_at);
      }
      try { 
        let p = JSON.parse(doc.content || '{}'); 
        if (Array.isArray(p) && p.length > 0) p = p[0];
        this.htmlContent = this.sanitizeImportedHtml(p.html) ?? '<div><br></div>'; 
        const el = document.querySelector('.page') as HTMLElement;
        if (el) {
          el.innerHTML = this.htmlContent;
          try { document.execCommand('enableObjectResizing', false, 'false'); } catch (e) { }
          setTimeout(() => {
             this.autoPaginate();
             this.updatePageHeight();
             
             // Re-paginate when images load to prevent them from intersecting pages
             const images = el.querySelectorAll('img');
             images.forEach(img => {
               if (!img.complete) {
                 img.addEventListener('load', () => {
                   this.autoPaginate();
                   this.updatePageHeight();
                 });
               }
             });

            // Setup selection change listener
            document.addEventListener('selectionchange', this.onSelectionChange.bind(this));
          }, 100);
        }
      } catch { }
      this.isLoadingDocument = false;
    });

    this.syncSub = this.api.connectSync(this.docId).subscribe(msg => {
      if (msg.type === 'reload_page') {
        window.location.reload();
      } else if (msg.type === 'presence') {
        this.activeUsers = msg.users ?? 1;
      } else if (msg.type === 'update') {
        const isInitialSync = msg.users === undefined;
        this.activeUsers = msg.users ?? this.activeUsers;
        
        if (isInitialSync && this.hasReceivedInitialSync) {
            // Ignore initial state dumps on reconnect to avoid overwriting local typing
            return;
        }
        if (isInitialSync) {
            this.hasReceivedInitialSync = true;
        }

        this.applyingRemote = true;
        if (msg.title) this.title = msg.title;
        if (msg.content !== undefined) {
          try {
            let p = JSON.parse(msg.content!);
            if (Array.isArray(p) && p.length > 0) p = p[0];
            const newHtml = this.sanitizeImportedHtml(p.html) ?? '<div><br></div>';
            if (newHtml !== this.htmlContent) {
              this.htmlContent = newHtml;
              const el = document.querySelector('.page') as HTMLElement;
              if (el) {
                el.innerHTML = this.htmlContent;
                try { document.execCommand('enableObjectResizing', false, 'false'); } catch (e) { }
                setTimeout(() => {
                   this.autoPaginate();
                   this.updatePageHeight();
                }, 100);
              }
            }
          } catch { }
        }
        setTimeout(() => this.applyingRemote = false, 50);
      }
    });
  }


  closeAllSubmenus() {
    document.querySelectorAll('.sub-dropdown').forEach(el => {
      (el as HTMLElement).style.transition = 'none'; // Force instant hide
      (el as HTMLElement).style.opacity = '0';
      (el as HTMLElement).style.visibility = 'hidden';
    });
  }

  positionSubmenu(event: MouseEvent, submenu: HTMLElement) {
    if (!submenu) return;
    
    // Clear inline styles (if previously hidden by scroll) for the currently hovered submenu
    submenu.style.transition = '';
    submenu.style.opacity = '';
    submenu.style.visibility = '';

    const parent = (event.currentTarget as HTMLElement);
    const parentRect = parent.getBoundingClientRect();
    
    submenu.style.position = 'fixed';
    
    // Overlap by 4px to prevent losing hover when moving mouse diagonally
    let leftPos = parentRect.right - 4;
    if (leftPos + 280 > window.innerWidth) {
      leftPos = parentRect.left - 280 + 4;
    }
    
    submenu.style.left = `${leftPos}px`;
    submenu.style.top = `${parentRect.top - 6}px`; // Shift up slightly to align text
    
    requestAnimationFrame(() => {
       const subRect = submenu.getBoundingClientRect();
       if (subRect.bottom > window.innerHeight) {
         let newTop = window.innerHeight - subRect.height - 10;
         submenu.style.top = `${Math.max(10, newTop)}px`;
       }
    });
  }

  onSelectionChange() {
    if (this.viewMode !== 'Editing') return;

    const selection = document.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    let node: Node | null = selection.anchorNode;
    let isInsideEditor = false;
    while (node) {
      if (node.nodeType === Node.ELEMENT_NODE && (node as Element).classList && (node as Element).classList.contains('page')) {
        isInsideEditor = true;
        break;
      }
      node = node.parentNode;
    }
    
    if (!isInsideEditor) return;

    // Update toolbar buttons based on selection
    this.isBold = document.queryCommandState('bold');
    this.isItalic = document.queryCommandState('italic');
    this.isUnderline = document.queryCommandState('underline');
    this.isStrikethrough = document.queryCommandState('strikeThrough');

    // Attempt to get font family
    const fontName = document.queryCommandValue('fontName');
    if (fontName) {
      this.activeFontFamily = fontName.replace(/["']/g, '').split(',')[0].trim();
    }

    // Attempt to get format block
    const formatBlock = document.queryCommandValue('formatBlock');
    if (formatBlock) {
      if (formatBlock === 'p' || formatBlock === 'div') this.activeBlockStyle = 'Normal text';
      else if (formatBlock === 'h1') this.activeBlockStyle = 'Title';
      else if (formatBlock === 'h2') this.activeBlockStyle = 'Subtitle';
      else if (formatBlock === 'h3') this.activeBlockStyle = 'Heading 1';
      else if (formatBlock === 'h4') this.activeBlockStyle = 'Heading 2';
      else if (formatBlock === 'h5') this.activeBlockStyle = 'Heading 3';
    }
  }

  toggleStar() {
    this.isStarred = !this.isStarred;
    this.showToast(this.isStarred ? 'Starred' : 'Unstarred');
  }

  toggleMenu(menu: string, event: Event) {
    event.stopPropagation();
    const target = event.target as HTMLElement;
    if (target && (target.closest('.dropdown') || target.closest('.sub-dropdown'))) {
      return;
    }
    this.activeMenu = this.activeMenu === menu ? null : menu;
  }

  closeMenus(event?: Event) {
    this.activeMenu = null;
    this.showFontSizeMenu = false;
    this.showContextMenu = false;
  }

  async onContextMenu(e: MouseEvent) {
    if (!this.activeSpellErrors) return;
    
    e.preventDefault();
    this.closeMenus();
    
    let word = '';
    let range: Range | null = null;
    if (document.caretRangeFromPoint) {
      range = document.caretRangeFromPoint(e.clientX, e.clientY);
      if (range && range.startContainer.nodeType === 3) {
         const text = range.startContainer.nodeValue || '';
         let start = range.startOffset;
         let end = range.startOffset;
         while (start > 0 && /\w/.test(text[start - 1])) start--;
         while (end < text.length && /\w/.test(text[end])) end++;
         word = text.substring(start, end);
         range.setStart(range.startContainer, start);
         range.setEnd(range.startContainer, end);
      }
    }

    this.contextMenuX = e.clientX;
    this.contextMenuY = e.clientY;
    this.showContextMenu = true;
    this.contextMenuSuggestions = [];
    this.contextMenuRange = range;

    if (word.trim().length > 1) {
       this.isFetchingSuggestions = true;
       try {
         const formData = new URLSearchParams();
         formData.append('language', 'en-US');
         formData.append('text', word);
         const res = await fetch('https://api.languagetool.org/v2/check', {
           method: 'POST',
           body: formData,
           headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
         });
         const data = await res.json();
         if (data.matches && data.matches.length > 0 && data.matches[0].replacements) {
            this.contextMenuSuggestions = data.matches[0].replacements.slice(0, 5).map((r: any) => r.value);
         }
       } catch (err) {
         console.error('Spellcheck error:', err);
       } finally {
         this.isFetchingSuggestions = false;
       }
    }
  }

  applySuggestion(suggestion: string) {
    if (this.contextMenuRange) {
      const sel = window.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(this.contextMenuRange);
      document.execCommand('insertText', false, suggestion);
    }
    this.showContextMenu = false;
  }

  newDoc() {
    this.api.createDocument('Untitled', 'doc').subscribe((doc: any) => {
      this.closeMenus();
      window.location.href = `/doc/${doc.id}`;
    });
  }

  loadTemplate(type: string) {
    let title = 'Untitled Document';
    let htmlContent = '';
    
    switch(type) {
      case 'resume':
        title = 'Resume';
        htmlContent = `<div style="padding:40px; font-family:Arial, sans-serif;">
          <h1 style="text-align:center; font-size: 32px; border-bottom: 2px solid #ccc; padding-bottom: 10px;">[Your Name]</h1>
          <p style="text-align:center;">[123 Address St, City, ST 12345] | [phone] | [email]</p>
          <br>
          <h2 style="color: #2c3e50;">Objective</h2>
          <p>[A brief objective statement...]</p>
          <h2 style="color: #2c3e50;">Experience</h2>
          <p><strong>[Job Title]</strong> at [Company Name] (YYYY-YYYY)</p>
          <ul><li>[Accomplishment 1]</li><li>[Accomplishment 2]</li></ul>
          <h2 style="color: #2c3e50;">Education</h2>
          <p><strong>[Degree]</strong> from [University Name] (YYYY)</p>
        </div>`;
        break;
      case 'letter':
        title = 'Business Letter';
        htmlContent = `<div style="padding:40px; font-family:Georgia, serif;">
          <p>[Date]</p><br>
          <p>[Recipient Name]<br>[Recipient Title]<br>[Company Name]<br>[Address]</p>
          <br><p>Dear [Recipient Name],</p>
          <p>I am writing to you today regarding...</p>
          <p>[Body of the letter...]</p>
          <br><p>Sincerely,</p>
          <p>[Your Name]</p>
        </div>`;
        break;
      case 'proposal':
        title = 'Project Proposal';
        htmlContent = `<div style="padding:40px; font-family:Arial, sans-serif;">
          <h1 style="text-align:center; color: #1a73e8; font-size: 36px;">Project Proposal</h1>
          <h3 style="text-align:center;">[Project Title]</h3>
          <p style="text-align:center; color: #555;">Prepared by: [Your Name]</p>
          <hr>
          <h2>1. Executive Summary</h2>
          <p>[Overview of the project...]</p>
          <h2>2. Goals & Objectives</h2>
          <ul><li>[Goal 1]</li><li>[Goal 2]</li></ul>
          <h2>3. Timeline & Milestones</h2>
          <table border="1" cellpadding="8" style="width:100%; border-collapse:collapse;">
            <tr><th style="background:#f4f4f9;">Phase</th><th style="background:#f4f4f9;">Deadline</th></tr>
            <tr><td>Phase 1: Research</td><td>[Date]</td></tr>
          </table>
        </div>`;
        break;
      case 'notes':
        title = 'Meeting Notes';
        htmlContent = `<div style="padding:40px;">
          <h1 style="border-bottom: 1px solid #ccc;">Meeting Notes</h1>
          <p><strong>Date:</strong> [Date] &nbsp;&nbsp;&nbsp; <strong>Time:</strong> [Time]</p>
          <p><strong>Attendees:</strong> [Name 1], [Name 2]</p>
          <h2>Agenda</h2>
          <ol><li>[Topic 1]</li><li>[Topic 2]</li></ol>
          <h2>Action Items</h2>
          <ul>
             <li><input type="checkbox"> [Task 1] (Assigned to: [Name])</li>
             <li><input type="checkbox"> [Task 2] (Assigned to: [Name])</li>
          </ul>
        </div>`;
        break;
      case 'brochure':
        title = 'Brochure';
        htmlContent = `<div style="padding:20px; font-family:'Trebuchet MS', sans-serif;">
          <h1 style="text-align:center; color: #e91e63; font-size: 40px;">[Product/Company Name]</h1>
          <p style="text-align:center; font-size: 18px; font-style: italic;">[Catchy Slogan Here]</p>
          <hr style="border: 0; height: 1px; background: #e91e63;">
          <div style="display:flex; gap: 20px; margin-top:20px;">
            <div style="flex:1; background:#f9f9f9; padding:15px; border-radius:8px;">
               <h3 style="color:#e91e63;">Feature 1</h3>
               <p>Discover our amazing new feature...</p>
            </div>
            <div style="flex:1; background:#f9f9f9; padding:15px; border-radius:8px;">
               <h3 style="color:#e91e63;">Feature 2</h3>
               <p>Unmatched quality and performance...</p>
            </div>
          </div>
        </div>`;
        break;
      case 'newsletter':
        title = 'Newsletter';
        htmlContent = `<div style="padding:20px; font-family:Arial, sans-serif; background:#f4f4f9; min-height: 100vh;">
           <div style="background:#4caf50; color:#fff; padding:30px; text-align:center; border-radius:8px;">
             <h1 style="margin:0;">The Weekly Update</h1>
             <p>[Month, Year] Edition</p>
           </div>
           <div style="padding:20px; background:#fff; margin-top:20px; border-radius:8px;">
             <h2>Top Story</h2>
             <p>This week we achieved a massive milestone...</p>
             <h2>Company News</h2>
             <ul><li>Update 1</li><li>Update 2</li></ul>
           </div>
        </div>`;
        break;
      case 'invoice':
        title = 'Invoice';
        htmlContent = `<div style="padding:40px; font-family:Arial, sans-serif;">
           <div style="display:flex; justify-content:space-between;">
             <div><h1 style="color:#2196f3;">INVOICE</h1><p>Invoice #: [1001]<br>Date: [Date]</p></div>
             <div style="text-align:right;"><h3>[Your Company]</h3><p>[Address]<br>[Email]</p></div>
           </div>
           <hr>
           <p><strong>Bill To:</strong><br>[Client Name]<br>[Client Address]</p>
           <table border="1" cellpadding="10" style="width:100%; border-collapse:collapse; margin-top:20px;">
             <tr style="background:#eee;"><th>Description</th><th>Qty</th><th>Price</th><th>Total</th></tr>
             <tr><td>[Service/Product]</td><td>1</td><td>$100.00</td><td>$100.00</td></tr>
             <tr><td colspan="3" style="text-align:right;"><strong>Subtotal</strong></td><td>$100.00</td></tr>
           </table>
        </div>`;
        break;
      case 'report':
        title = 'Business Report';
        htmlContent = `<div style="padding:40px;">
           <h1 style="text-align:center; font-size: 32px; color: #3f51b5;">Business Report</h1>
           <h3 style="text-align:center;">Q3 Financial Review</h3>
           <p style="text-align:center;">Prepared by [Name] - [Date]</p>
           <br><br>
           <h2>1. Introduction</h2>
           <p>This report details the financial performance for...</p>
           <h2>2. Key Findings</h2>
           <p>Overall revenue grew by 15% compared to...</p>
           <h2>3. Conclusion</h2>
           <p>The strategic initiatives implemented have proven successful...</p>
        </div>`;
        break;
      case 'plan':
        title = 'Business Plan';
        htmlContent = `<div style="padding:40px;">
           <h1 style="text-align:center;">Business Plan</h1>
           <h3 style="text-align:center; color:#555;">[Startup Name]</h3>
           <br>
           <h2>1. Executive Summary</h2><p>[Summary]</p>
           <h2>2. Market Analysis</h2><p>[Analysis]</p>
           <h2>3. Marketing Strategy</h2><p>[Strategy]</p>
           <h2>4. Financial Projections</h2><p>[Projections]</p>
        </div>`;
        break;
      case 'essay':
        title = 'Academic Essay';
        htmlContent = `<div style="padding:40px; font-family:'Times New Roman', Times, serif; line-height:2;">
           <div style="text-align:left;">[Your Name]<br>[Professor's Name]<br>[Course Name]<br>[Date]</div>
           <h2 style="text-align:center; margin-top:30px;">[Catchy Essay Title]</h2>
           <p style="text-indent: 40px;">[The first paragraph of your essay begins here. Remember to include your thesis statement...]</p>
           <p style="text-indent: 40px;">[The next body paragraph starts here...]</p>
        </div>`;
        break;
    }

    this.api.createDocument(title, 'doc').subscribe((doc: any) => {
      this.api.saveDocument(doc.id, title, JSON.stringify({ html: htmlContent })).subscribe(() => {
        this.closeMenus();
        window.location.href = `/doc/${doc.id}`;
      });
    });
  }

  renameDoc() {
    this.closeMenus();
    setTimeout(() => {
      const el = document.querySelector('.title-input') as HTMLInputElement;
      if (el) el.focus();
    }, 10);
  }

  makeCopy() {
    this.closeMenus();
    this.api.createDocument(this.title + ' (Copy)', 'doc').subscribe((doc: any) => {
      this.api.saveDocument(doc.id, doc.title, JSON.stringify({ html: this.htmlContent })).subscribe(() => {
        window.location.href = `/doc/${doc.id}`;
      });
    });
  }

  trashDoc() {
    this.closeMenus();
    this.trashModalOpen = true;
  }

  confirmTrashDoc() {
    this.trashModalOpen = false;
    this.api.deleteDocument(this.docId).subscribe(() => {
      this.router.navigate(['/']);
    });
  }

  openFind() {
    this.closeMenus();
    this.findDialogVisible = true;
    setTimeout(() => {
      const input = document.querySelector('.find-dialog input') as HTMLInputElement;
      if (input) input.focus();
    }, 10);
  }

  executeFind() {
    if (this.findTextQuery) {
      const el = document.querySelector('.page') as HTMLElement;
      if (el) el.focus();
      
      // Try to find the text. If we hit the end, reset selection and search from top.
      if (!(window as any).find(this.findTextQuery)) {
        window.getSelection()?.removeAllRanges();
        if (!(window as any).find(this.findTextQuery)) {
          this.showToast('Text not found.');
        }
      }
    }
  }

  toggleFullScreen() {
    this.closeMenus();
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(() => {
        this.showToast('Could not enable full screen.');
      });
    } else {
      document.exitFullscreen();
    }
  }
  showFontSizeMenu = false;

  hideFontSizeMenu() {
    setTimeout(() => { this.showFontSizeMenu = false; }, 150);
  }

  toggleViewMode() { this.viewMode = this.viewMode === 'Editing' ? 'Viewing' : 'Editing'; this.closeMenus(); }
  togglePrintLayout() { this.showPrintLayout = !this.showPrintLayout; this.closeMenus(); }
  toggleRuler() { this.showRuler = !this.showRuler; this.closeMenus(); }
  toggleEquationToolbar() { this.showEquationToolbar = !this.showEquationToolbar; this.closeMenus(); }
  toggleNonPrinting() { this.showNonPrinting = !this.showNonPrinting; this.closeMenus(); }

  setDocumentView(view: string) {
    this.documentViewType = view;
    this.closeMenus();
    if (view === 'Web View') {
      const page = document.querySelector('.page') as HTMLElement;
      if (page) {
        page.style.width = '100%';
        page.style.maxWidth = '100%';
        page.style.minHeight = '100vh';
      }
    } else {
      const page = document.querySelector('.page') as HTMLElement;
      if (page) {
        page.style.width = '21cm';
        page.style.maxWidth = '';
        page.style.minHeight = '29.7cm';
      }
    }
  }

  setZoom(level: number) {
    this.zoomLevel = level;
    this.closeMenus();
  }

  fitWidth() {
    this.setZoom(150);
  }

  fitPageToWindow() {
    this.setZoom(100);
  }

  toggleReaderView() {
    this.isReaderView = !this.isReaderView;
    this.closeMenus();
  }

  toggleNavigator() {
    this.showNavigator = !this.showNavigator;
    if (this.showNavigator) {
       this.showToast('Navigator Enabled');
    }
    this.closeMenus();
  }

  toggleHideImages() {
    this.hideImages = !this.hideImages;
    this.closeMenus();
  }

  toggleBookmarks() {
    this.showBookmarks = !this.showBookmarks;
    this.closeMenus();
  }

  toggleSmartGridLines() {
    this.showSmartGridLines = !this.showSmartGridLines;
    this.closeMenus();
  }

  toggleObjectIndicator() {
    this.showObjectIndicator = !this.showObjectIndicator;
    this.closeMenus();
  }

  toggleFormatting() {
    this.showToast('Formatting Symbols Toggled');
    this.closeMenus();
  }

  setAppearance(mode: string) {
    this.appearanceMode = mode;
    this.closeMenus();
  }



  insertImage() {
    this.closeMenus();
    const url = prompt('Enter image URL:', 'https://');
    if (url) {
      this.insertImageAsBlock(url);
    }
  }

  onImageUpload(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const dataUrl = e.target?.result as string;
        this.insertImageAsBlock(dataUrl);
      };
      reader.readAsDataURL(file);
    }
    (event.target as HTMLInputElement).value = '';
  }

  onVideoUpload(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const url = URL.createObjectURL(file);
      const html = `<div contenteditable="false" style="display: inline-block; margin: 10px 0;"><video controls src="${url}" style="width: 100%; max-width: 500px;"></video></div><br>`;
      document.execCommand('insertHTML', false, html);
    }
    (event.target as HTMLInputElement).value = '';
  }

  onAudioUpload(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const url = URL.createObjectURL(file);
      const html = `<div contenteditable="false" style="display: inline-block; margin: 10px 0;"><audio controls src="${url}" style="width: 100%; max-width: 300px;"></audio></div><br>`;
      document.execCommand('insertHTML', false, html);
    }
    (event.target as HTMLInputElement).value = '';
  }

  onSignatureUpload(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const dataUrl = e.target?.result as string;
        const html = `<div contenteditable="false" style="display: inline-block; border-bottom: 2px solid #202124; padding: 0 16px; margin: 16px 0;"><img src="${dataUrl}" style="max-height: 50px; vertical-align: middle;" /></div><br>`;
        document.execCommand('insertHTML', false, html);
      };
      reader.readAsDataURL(file);
    }
    (event.target as HTMLInputElement).value = '';
  }

  onDocOpen(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target?.result as string;
        const content = file.name.endsWith('.html') ? text : text.replace(/\n/g, '<br>');
        
        this.api.createDocument(file.name.replace(/\.[^/.]+$/, ""), 'doc').subscribe((doc: any) => {
          this.api.saveDocument(doc.id, doc.title, JSON.stringify({ html: content })).subscribe(() => {
            window.location.href = `/doc/${doc.id}`;
          });
        });
      };
      reader.readAsText(file);
    }
    (event.target as HTMLInputElement).value = '';
  }

  emailDoc() {
    this.closeMenus();
    const subject = encodeURIComponent(`Document: ${this.title}`);
    const body = encodeURIComponent(`Check out this document I am sharing with you:\n\n${window.location.href}`);
    window.location.href = `mailto:?subject=${subject}&body=${body}`;
  }

  showVersionHistory() {
    this.closeMenus();
    this.versionModalOpen = true;
  }

  makeOffline() {
    this.closeMenus();
    try {
      localStorage.setItem(`offline_doc_${this.docId}`, this.htmlContent);
      this.showToast('Document cached for offline access.');
    } catch {
      this.showToast('Failed to enable offline access.');
    }
  }

  showDetails() {
    this.closeMenus();
    const el = document.querySelector('.page') as HTMLElement;
    const text = el ? (el.innerText || '') : '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    this.detailsData = { words, chars };
    this.detailsModalOpen = true;
  }

  showWordCount() {
    this.closeMenus();
    const el = document.querySelector('.page') as HTMLElement;
    if (el) {
      const text = el.innerText || '';
      this.wcWords = text.trim() ? text.trim().split(/\s+/).length : 0;
      this.wcChars = text.length;
      this.wordCountModalOpen = true;
    }
  }

  printDoc() {
    this.closeMenus();
    // Use @media print rules Gï¿½ï¿½ UI chrome is hidden via CSS
    window.print();
  }

  getTodayDate(): string {
    const d = new Date();
    return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
  }

  insertPhoneField() {
    this.api.getProfile().subscribe({
      next: (profile: any) => {
        const userPhone = profile?.phone || (this.auth.user as any)?.phone || '';
        this.insertField(userPhone);
      },
      error: () => {
        const fallbackPhone = (this.auth.user as any)?.phone || '';
        this.insertField(fallbackPhone);
      }
    });
  }

  insertField(text: string) {
    this.closeMenus();
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      document.execCommand('insertText', false, text);
    }
  }

  async insertLink() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const url = await this.showPrompt('Enter link URL:', 'https://');
    if (url && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      document.execCommand('createLink', false, url);
    }
  }

  openImageModal(type: string) {
    this.closeMenus();
    this.imageModalType = type;
    this.imageUrlInput = '';
    this.imageModalVisible = true;
  }

  confirmImageUrl() {
    if (this.imageUrlInput) {
      this.insertImageAsBlock(this.imageUrlInput);
    }
    this.imageModalVisible = false;
  }

  setLineSpacing(val: string) {
    this.closeMenus();
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
       let node = selection.anchorNode;
       if (node?.nodeType === Node.TEXT_NODE) node = node.parentNode;
       let block = node as HTMLElement;
       while (block && block.tagName !== 'DIV' && block.tagName !== 'P' && !block.classList?.contains('page')) {
           block = block.parentElement as HTMLElement;
       }
       if (block && !block.classList?.contains('page')) {
           block.style.lineHeight = val;
       } else {
           document.execCommand('formatBlock', false, 'DIV');
           let newNode = selection.anchorNode;
           if (newNode?.nodeType === Node.TEXT_NODE) newNode = newNode.parentNode;
           if (newNode && (newNode as HTMLElement).tagName === 'DIV') {
               (newNode as HTMLElement).style.lineHeight = val;
           }
       }
    }
    this.save();
  }

  exec(cmd: string) { 
    this.closeMenus();
    const el = document.querySelector('.page') as HTMLElement;
    if (el) {
      if (!el.contains(window.getSelection()?.anchorNode || null)) {
        el.focus();
      }
      try { document.execCommand('enableObjectResizing', false, 'false'); } catch (e) { }
    }
    document.execCommand(cmd, false); 
    setTimeout(() => {
      this.updateOverlay();
      this.updatePageHeight();
      this.save();
    }, 10);
  }

  execVal(cmd: string, val: string | Event) {
    this.closeMenus();
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();

    let value = typeof val === 'string' ? val : (val.target as HTMLSelectElement).value;
    if (value) document.execCommand(cmd, false, value);

    if (typeof val !== 'string') {
      (val.target as HTMLSelectElement).value = '';
    }

    setTimeout(() => {
      this.updateOverlay();
      this.updatePageHeight();
      this.save();
    }, 10);
  }

  changeCase(type: string) {
    this.closeMenus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const text = selection.toString();
    if (!text) return;
    
    let newText = '';
    switch(type) {
      case 'uppercase': newText = text.toUpperCase(); break;
      case 'lowercase': newText = text.toLowerCase(); break;
      case 'capitalize': newText = text.toLowerCase().replace(/\b\w/g, l => l.toUpperCase()); break;
      case 'sentence': newText = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase(); break;
      case 'smallcaps': 
         document.execCommand('insertHTML', false, `<span style="font-variant: small-caps;">${text}</span>`);
         return;
    }
    if (newText) document.execCommand('insertText', false, newText);
  }

  setCharSpacing(pt: string) {
    this.activeCharSpacing = pt;
    this.closeMenus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const text = selection.toString();
    if (text) {
      document.execCommand('insertHTML', false, `<span style="letter-spacing: ${pt === 'Normal' ? 'normal' : pt};">${text}</span>`);
    } else {
      this.showToast(`Character spacing set to ${pt}`);
    }
  }

  execDir(dir: string) {
    this.activeDir = dir;
    this.closeMenus();
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    const selection = window.getSelection();
    if (selection && selection.anchorNode) {
      let block = selection.anchorNode as HTMLElement;
      if (block.nodeType === 3) block = block.parentElement as HTMLElement;
      let target = block.closest('p, div, h1, h2, h3, h4, h5, h6, li');
      if (target) {
        (target as HTMLElement).setAttribute('dir', dir);
      } else {
        document.execCommand('insertHTML', false, `<div dir="${dir}">${selection.toString() || '&#8203;'}</div>`);
      }
    }
  }

  applyDropCap(style: string) {
    this.closeMenus();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    
    let node: any = sel.anchorNode;
    const page = document.querySelector('.page');
    while (node && node.parentNode !== page && node !== page) {
      node = node.parentNode;
    }
    if (!node || node === page) return;

    let dropCap = (node as HTMLElement).querySelector ? (node as HTMLElement).querySelector('.drop-cap') as HTMLElement : null;
    
    if (style === 'none') {
        if (dropCap) {
           const text = document.createTextNode(dropCap.innerText);
           dropCap.parentNode?.replaceChild(text, dropCap);
        }
        return;
    }
    
    if (!dropCap) {
       const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, null);
       const firstText = walker.nextNode();
       if (firstText && firstText.nodeValue && firstText.nodeValue.trim().length > 0) {
           const val = firstText.nodeValue;
           let i = 0;
           while (i < val.length && val[i].trim() === '') i++;
           if (i < val.length) {
               const char = val[i];
               const span = document.createElement('span');
               span.className = 'drop-cap';
               span.innerText = char;
               const afterText = document.createTextNode(val.substring(i + 1));
               const beforeText = document.createTextNode(val.substring(0, i));
               const parent = firstText.parentNode;
               parent?.insertBefore(beforeText, firstText);
               parent?.insertBefore(span, firstText);
               parent?.insertBefore(afterText, firstText);
               parent?.removeChild(firstText);
               dropCap = span;
           }
       }
    }
    
    if (dropCap) {
       let css = 'float: left; font-size: 3em; line-height: 0.8; margin-right: 8px; font-weight: bold; padding: 4px;';
       if (style === 'style2') css += ' background: #8ab4f8; color: white; border-radius: 4px;';
       if (style === 'style3') css += ' background: #81c995; color: white; border-radius: 12px;';
       if (style === 'style4') css += ' border-bottom: 3px solid #8ab4f8;';
       if (style === 'style5') css += ' background: #f48fb1; color: white; border-radius: 4px 24px 24px 4px;';
       dropCap.style.cssText = css;
    }
  }

  applyBlockQuote(style: string) {
    this.closeMenus();
    document.execCommand('formatBlock', false, 'blockquote');
    
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    let node: any = sel.anchorNode;
    while (node && node !== document.querySelector('.page') && node.nodeName !== 'BLOCKQUOTE') {
      node = node.parentNode;
    }
    
    if (node && node.nodeName === 'BLOCKQUOTE') {
       let css = '';
       if (style === 'style1') css = 'border-left: 4px solid #1a73e8; background: #e8f0fe; padding: 12px; margin: 16px 0; border-radius: 4px;';
       if (style === 'style2') css = 'border-radius: 4px; background: #fce8e6; padding: 12px; margin: 16px 0;';
       if (style === 'style3') css = 'border-radius: 4px; background: #fef7e0; padding: 12px; margin: 16px 0;';
       if (style === 'style4') css = 'border-radius: 4px; background: #f3f3f3; padding: 12px; margin: 16px 0;';
       if (style === 'style5') css = 'border-radius: 4px; background: #3c4043; color: white; padding: 12px; margin: 16px 0;';
       if (style === 'style6') css = 'border-left: 4px solid #00bcd4; background: #e0f7fa; padding: 12px; margin: 16px 0; border-radius: 4px;';
       (node as HTMLElement).style.cssText = css;
    }
  }
  changeFont(font: string) {
    this.activeFontFamily = font;
    document.execCommand('fontName', false, font);
    this.closeMenus();
    setTimeout(() => { this.updatePageHeight(); this.save(); }, 10);
  }

  changeFontSize(size: number) {
    this.activeFontSize = size;
    const val = size + 'pt';
    const el = document.querySelector('.page') as HTMLElement;
    if (!el) return;
    el.focus();

    const sel = window.getSelection();
    if (sel && sel.isCollapsed) {
      const span = document.createElement('span');
      span.style.fontSize = val;
      span.innerHTML = '&#8203;';
      const range = sel.getRangeAt(0);
      range.insertNode(span);
      range.setStart(span.firstChild!, 1);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
    document.execCommand('fontSize', false, '7'); 
    const fontEls = document.querySelectorAll('font[size="7"]');
    fontEls.forEach(f => {
        f.removeAttribute('size');
        f.setAttribute('style', `font-size: ${val}`);
    });
    }

    this.closeMenus();
    setTimeout(() => { this.updatePageHeight(); this.save(); }, 10);
  }

  onFontSizeInputChange() {
    if (this.activeFontSize > 0) this.changeFontSize(this.activeFontSize);
  }

  incrementFontSize() {
    this.changeFontSize(this.activeFontSize + 1);
  }

  decrementFontSize() {
    if (this.activeFontSize > 1) {
      this.changeFontSize(this.activeFontSize - 1);
    }
  }

  changeStyle(style: string, label: string) {
    this.activeBlockStyle = label;
    document.execCommand('formatBlock', false, style);
    this.closeMenus();
    setTimeout(() => { this.updatePageHeight(); this.save(); }, 10);
  }

  insertText(text: string) {
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    document.execCommand('insertText', false, text);
    this.closeMenus();
  }

  insertImageAsBlock(dataUrl: string) {
    const el = document.querySelector('.page') as HTMLElement;
    this.closeMenus();
    el.focus();
    
    // Compress image to prevent WebSocket limits from disconnecting the client
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;
      const MAX_DIM = 1200;
      if (width > MAX_DIM || height > MAX_DIM) {
        const ratio = Math.min(MAX_DIM / width, MAX_DIM / height);
        width *= ratio;
        height *= ratio;
      }
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (ctx) {
         ctx.drawImage(img, 0, 0, width, height);
         const compressedDataUrl = canvas.toDataURL('image/jpeg', 0.8);
         const html = `<div><img src="${compressedDataUrl}" style="max-width: 100%; height: auto; display: block; margin: 0 auto; cursor: pointer;"></div><br>`;
         document.execCommand('insertHTML', false, html);
         this.save();
      } else {
         const html = `<div><img src="${dataUrl}" style="max-width: 100%; height: auto; display: block; margin: 0 auto; cursor: pointer;"></div><br>`;
         document.execCommand('insertHTML', false, html);
         this.save();
      }
    };
    img.onerror = () => {
      // Fallback if image fails to load for compression
      const html = `<div><img src="${dataUrl}" style="max-width: 100%; height: auto; display: block; margin: 0 auto; cursor: pointer;"></div><br>`;
      document.execCommand('insertHTML', false, html);
      this.save();
    };
    img.src = dataUrl;
  }

  insertHTML(html: string) {
    const el = document.querySelector('.page') as HTMLElement;
    if (!el) return;

    el.focus();
    const sel = window.getSelection();
    if (sel) {
      let hasValidRange = false;
      if (sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        if (el.contains(range.commonAncestorContainer)) {
          hasValidRange = true;
        }
      }

      // If no valid selection inside the editor, place cursor at the end
      if (!hasValidRange) {
        const range = document.createRange();
        range.selectNodeContents(el);
        range.collapse(false); // collapse to end
        sel.removeAllRanges();
        sel.addRange(range);
      }
    }

    document.execCommand('insertHTML', false, html);
    this.closeMenus();
  }

  async insertTable() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const input = await this.showPrompt('Enter table dimensions (Rows x Columns), e.g., "3x4":', '3x3');
    if (!input || !savedRange) return;
    sel?.removeAllRanges();
    sel?.addRange(savedRange);
    
    const parts = input.toLowerCase().split('x');
    const rows = parseInt(parts[0], 10) || 3;
    const cols = parseInt(parts[1], 10) || 3;
    
    let html = `<table style="width: 100%; border-collapse: collapse; margin-bottom: 1rem;">`;
    for (let r = 0; r < rows; r++) {
      html += `<tr>`;
      for (let c = 0; c < cols; c++) {
        html += `<td style="border: 1px solid #dadce0; padding: 8px; min-width: 50px;"></td>`;
      }
      html += `</tr>`;
    }
    html += `</table><br>`;
    
    document.execCommand('insertHTML', false, html);
  }

  insertBreak() {
    const el = document.querySelector('.page') as HTMLElement;
    if (!el) return;

    // Get current selection
    const sel = window.getSelection();
    let targetBlock: HTMLElement | null = null;

    if (sel && sel.rangeCount > 0) {
      let node = sel.getRangeAt(0).commonAncestorContainer as HTMLElement;
      if (node.nodeType === Node.TEXT_NODE) node = node.parentElement as HTMLElement;

      // Traverse up to find the direct child of .page
      while (node && node !== el && node.parentElement !== el) {
        node = node.parentElement as HTMLElement;
      }
      if (node && node.parentElement === el) {
        targetBlock = node as HTMLElement;
      }
    }

    const breakHTML = '<div class="manual-page-break" style="page-break-after: always;"></div><div class="page-break-content" style="min-height: 20px;">&#8203;</div>';

    if (targetBlock) {
      targetBlock.insertAdjacentHTML('afterend', breakHTML);
      const newBlock = targetBlock.nextElementSibling?.nextElementSibling;
      if (newBlock) {
        const range = document.createRange();
        range.selectNodeContents(newBlock);
        range.collapse(false);
        sel?.removeAllRanges();
        sel?.addRange(range);
      }
    } else {
      el.insertAdjacentHTML('beforeend', breakHTML);
      const newBlock = el.lastElementChild;
      if (newBlock) {
        const range = document.createRange();
        range.selectNodeContents(newBlock);
        range.collapse(false);
        sel?.removeAllRanges();
        sel?.addRange(range);
      }
    }

    this.closeMenus();
    this.onInput(el); // Force update
  }

  insertBuildingBlock() {
    this.insertHTML('<strong>Meeting Notes</strong><br><em>Date:</em><br><em>Attendees:</em><br><em>Action Items:</em><br><ul><li><br></li></ul><br>');
  }

  async insertSmartChip() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const name = await this.showPrompt('Enter user name for Smart Chip:');
    if (name && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const el = document.querySelector('.page') as HTMLElement;
      const html = `<span contenteditable="false" style="background:#e8f0fe; color:#1a73e8; padding:2px 8px; border-radius:12px; font-size:12px; font-weight:500; cursor:pointer;" onclick="alert('Profile: ${name}')">@${name}</span> `;
      document.execCommand('insertHTML', false, html);
    }
  }

  async insertSignature() {
    this.closeMenus();
    const sel = window.getSelection();
    let savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    const name = await this.showPrompt('Enter your name for eSignature:');
    if (name && savedRange) {
      sel?.removeAllRanges();
      sel?.addRange(savedRange);
      const el = document.querySelector('.page') as HTMLElement;
      const html = `<div contenteditable="false" style="display:inline-block; border-bottom: 2px solid #202124; padding:0 16px; margin: 16px 0; font-family: 'Caveat', cursive; font-size: 24px; color: #1a73e8;">${name}</div><br>`;
      document.execCommand('insertHTML', false, html);
    }
  }

  insertDrawing() {
    this.closeMenus();
    this.drawingModalVisible = true;
    setTimeout(() => {
      const canvas = this.drawCanvasRef?.nativeElement;
      if (canvas) {
        this.drawCtx = canvas.getContext('2d');
        if (this.drawCtx) {
          this.drawCtx.fillStyle = '#ffffff';
          this.drawCtx.fillRect(0, 0, canvas.width, canvas.height);
          this.drawCtx.strokeStyle = '#1a73e8';
          this.drawCtx.lineWidth = 3;
          this.drawCtx.lineCap = 'round';
        }
      }
    }, 100);
  }

  startDraw(e: MouseEvent) {
    this.isDrawing = true;
    this.lastX = e.offsetX;
    this.lastY = e.offsetY;
  }
  
  draw(e: MouseEvent) {
    if (!this.isDrawing || !this.drawCtx) return;
    this.drawCtx.beginPath();
    this.drawCtx.moveTo(this.lastX, this.lastY);
    this.drawCtx.lineTo(e.offsetX, e.offsetY);
    this.drawCtx.stroke();
    this.lastX = e.offsetX;
    this.lastY = e.offsetY;
  }
  
  stopDraw() { this.isDrawing = false; }
  
  clearDrawing() {
    const canvas = this.drawCanvasRef?.nativeElement;
    if (this.drawCtx && canvas) {
      this.drawCtx.fillStyle = '#ffffff';
      this.drawCtx.fillRect(0, 0, canvas.width, canvas.height);
    }
  }

  insertDrawingImage() {
    const canvas = this.drawCanvasRef?.nativeElement;
    if (canvas) {
      const dataUrl = canvas.toDataURL('image/png');
      this.drawingModalVisible = false;
      const el = document.querySelector('.page') as HTMLElement;
      if (el) el.focus();
      this.insertImageAsBlock(dataUrl);
    }
  }

  insertChart() {
    this.closeMenus();
    this.chartModalVisible = true;
  }

  openChartModal(type: string = 'Bar') {
    this.closeMenus();
    this.chartConfig.type = type.replace(' Chart', '').replace('Stacked ', '');
    this.chartModalVisible = true;
  }

  triggerExcelImport() {
    this.showToast('Importing from Excel...');
    setTimeout(() => {
      this.chartTitle = 'Excel Imported Data';
      this.chartConfig.xAxisLabels = 'Jan, Feb, Mar, Apr, May';
      this.chartValues = '15, 45, 30, 90, 60';
      if (this.chartConfig.showZ) {
        this.chartConfig.zAxisValues = '5, 20, 10, 40, 25';
      }
      this.showToast('Excel Data Imported Successfully');
    }, 1000);
  }

  generateAndInsertChart() {
    const yVals = this.chartValues.split(',').map(v => parseInt(v.trim(), 10)).filter(v => !isNaN(v));
    const labels = this.chartConfig.xAxisLabels.split(',').map(v => v.trim());
    const zVals = this.chartConfig.showZ ? this.chartConfig.zAxisValues.split(',').map(v => parseInt(v.trim(), 10)).filter(v => !isNaN(v)) : [];
    
    if (yVals.length === 0) {
      alert('Please enter valid numbers for Y-Axis.');
      return;
    }
    
    const width = 500;
    const height = 300;
    const baseColor = this.chartConfig.color;
    let svg = `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg" style="background:#fff; border:1px solid #ccc; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">`;
    svg += `<text x="${width / 2}" y="30" font-family="Arial" font-size="16" font-weight="bold" text-anchor="middle" fill="#202124">${this.chartTitle}</text>`;
    
    const maxVal = Math.max(...yVals, ...zVals, 1);
    const plotH = height - 80;
    const plotW = width - 80;
    const originX = 40;
    const originY = height - 40;
    
    if (this.chartConfig.type !== 'Pie' && this.chartConfig.type !== 'Donut') {
      svg += `<line x1="${originX}" y1="${originY}" x2="${originX + plotW}" y2="${originY}" stroke="#e0e0e0" stroke-width="2"/>`;
      svg += `<line x1="${originX}" y1="${originY}" x2="${originX}" y2="${originY - plotH}" stroke="#e0e0e0" stroke-width="2"/>`;
      
      // Draw Y-axis ticks and horizontal grid lines
      const numTicks = 5;
      for (let i = 0; i <= numTicks; i++) {
        const tickVal = Math.round((maxVal / numTicks) * i);
        const yPos = originY - (tickVal / maxVal) * plotH;
        svg += `<text x="${originX - 10}" y="${yPos + 4}" font-family="Arial" font-size="11" text-anchor="end" fill="#5f6368">${tickVal}</text>`;
        if (i > 0) {
          svg += `<line x1="${originX}" y1="${yPos}" x2="${originX + plotW}" y2="${yPos}" stroke="#f1f3f4" stroke-width="1"/>`;
        }
      }
    }

    const stepX = plotW / Math.max(yVals.length, 1);

    if (this.chartConfig.type === 'Pie' || this.chartConfig.type === 'Donut') {
      const cx = width / 2;
      const cy = height / 2 + 10;
      const r = Math.min(plotW, plotH) / 2;
      const total = yVals.reduce((a, b) => a + b, 0);
      let currentAngle = -Math.PI / 2;
      const pieColors = [baseColor, '#ea4335', '#fbbc04', '#34a853', '#673ab7', '#ff9800'];
      
      if (total === 0) {
        svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#e0e0e0" />`;
      } else {
        yVals.forEach((val, i) => {
          const sliceAngle = (val / total) * 2 * Math.PI;
          const x1 = cx + r * Math.cos(currentAngle);
          const y1 = cy + r * Math.sin(currentAngle);
          currentAngle += sliceAngle;
          const x2 = cx + r * Math.cos(currentAngle);
          const y2 = cy + r * Math.sin(currentAngle);
          const largeArc = sliceAngle > Math.PI ? 1 : 0;
          
          if (val === total) {
            svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${pieColors[i % pieColors.length]}" />`;
          } else {
            svg += `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${pieColors[i % pieColors.length]}" stroke="#fff" stroke-width="1"/>`;
          }
          
          const lblAngle = currentAngle - sliceAngle / 2;
          const lblX = cx + (r * 0.7) * Math.cos(lblAngle);
          const lblY = cy + (r * 0.7) * Math.sin(lblAngle);
          svg += `<text x="${lblX}" y="${lblY}" font-family="Arial" font-size="12" font-weight="bold" text-anchor="middle" fill="#fff">${labels[i] || val}</text>`;
        });
      }
      if (this.chartConfig.type === 'Donut') {
        svg += `<circle cx="${cx}" cy="${cy}" r="${r * 0.5}" fill="#fff" />`;
      }
    } else if (this.chartConfig.type === 'Line') {
      let pathD = '';
      yVals.forEach((v, i) => {
        const x = originX + i * stepX + stepX / 2;
        const y = originY - (v / maxVal) * plotH;
        pathD += `${i === 0 ? 'M' : 'L'} ${x} ${y} `;
        svg += `<circle cx="${x}" cy="${y}" r="4" fill="${baseColor}" />`;
        svg += `<text x="${x}" y="${originY + 20}" font-family="Arial" font-size="11" text-anchor="middle" fill="#5f6368">${labels[i] || ''}</text>`;
      });
      if (pathD) svg += `<path d="${pathD}" fill="none" stroke="${baseColor}" stroke-width="3" />`;
      
      if (this.chartConfig.showZ && zVals.length) {
        let zPath = '';
        const zColor = '#ea4335';
        zVals.forEach((v, i) => {
          const x = originX + i * stepX + stepX / 2;
          const y = originY - (v / maxVal) * plotH;
          zPath += `${i === 0 ? 'M' : 'L'} ${x} ${y} `;
          svg += `<circle cx="${x}" cy="${y}" r="4" fill="${zColor}" />`;
        });
        if (zPath) svg += `<path d="${zPath}" fill="none" stroke="${zColor}" stroke-width="3" />`;
      }
    } else {
      const barW = (stepX * 0.6) / (this.chartConfig.showZ ? 2 : 1);
      yVals.forEach((v, i) => {
        const h = (v / maxVal) * plotH;
        const x = originX + i * stepX + stepX / 2 - (this.chartConfig.showZ ? barW : barW / 2);
        const y = originY - h;
        svg += `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${baseColor}" rx="2" />`;
        svg += `<text x="${originX + i * stepX + stepX / 2}" y="${originY + 20}" font-family="Arial" font-size="11" text-anchor="middle" fill="#5f6368">${labels[i] || ''}</text>`;
      });
      if (this.chartConfig.showZ && zVals.length) {
        const zColor = '#ea4335';
        zVals.forEach((v, i) => {
          const h = (v / maxVal) * plotH;
          const x = originX + i * stepX + stepX / 2;
          const y = originY - h;
          svg += `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${zColor}" rx="2" />`;
        });
      }
    }

    svg += `</svg>`;
    const svg64 = btoa(unescape(encodeURIComponent(svg)));
    const image64 = 'data:image/svg+xml;base64,' + svg64;

    this.chartModalVisible = false;
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    this.insertImageAsBlock(image64);
  }


  insertTOC(type: string) {
    this.closeMenus();
    const isLinks = type === 'links';
    const html = `
      <div style="border: 1px dashed #ccc; padding: 12px; margin: 12px 0; background: #fafafa; font-family: sans-serif;">
        <h3 style="margin-top: 0;">Table of Contents</h3>
        <ul style="padding-left: 20px;">
          <li>${isLinks ? '<a href="#section1">' : ''}Section 1: Introduction${isLinks ? '</a>' : ''}</li>
          <li>${isLinks ? '<a href="#section2">' : ''}Section 2: Methods${isLinks ? '</a>' : ''}</li>
          <li>${isLinks ? '<a href="#section3">' : ''}Section 3: Results${isLinks ? '</a>' : ''}</li>
        </ul>
      </div><br>
    `;
    const el = document.querySelector('.page') as HTMLElement;
    if (el) el.focus();
    document.execCommand('insertHTML', false, html);
  }

  onInput(editor: HTMLElement) {
    if (this.applyingRemote) return;
    
    // Check for newly pasted/inserted images that need to load before we can measure them properly
    const newImages = Array.from(editor.querySelectorAll('img:not([data-loaded="true"])')) as HTMLImageElement[];
    if (newImages.length > 0) {
       newImages.forEach(img => {
           // Mark as tracked so we don't attach multiple listeners
           img.dataset['loaded'] = 'true';
           if (!img.complete) {
               const onLoadOrError = () => {
                   if (this.typingTimer) clearTimeout(this.typingTimer);
                   this.typingTimer = window.setTimeout(() => this.save(), 400);
               };
               img.addEventListener('load', onLoadOrError, { once: true });
               img.addEventListener('error', onLoadOrError, { once: true });
           }
       });
    }

    // Run pagination synchronously to prevent text from overflowing visually while typing
    this.autoPaginate();

    // For small docs, we can update immediately, but for large docs it's best to debounce
    if (this.inputTimeout) {
      clearTimeout(this.inputTimeout);
    }
    
    this.inputTimeout = setTimeout(() => {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = editor.innerHTML;
      tempDiv.querySelectorAll('.page-break-dummy').forEach(row => row.remove());
      this.htmlContent = tempDiv.innerHTML;
      
      this.updateOverlay();
      this.updatePageHeight();
      this.updateCounts();
      this.api.sendUpdate(JSON.stringify({ html: this.htmlContent }), this.title, this.autoSaveEnabled);
    }, 400); // 400ms debounce ensures smooth typing without locking the main thread
  }

  autoPaginate() {
    const pageEl = document.querySelector('.page') as HTMLElement;
    if (!pageEl || !this.showPrintLayout) return;
    
    // Auto-wrap raw text nodes that the browser creates to ensure we can measure them
    const childNodes = Array.from(pageEl.childNodes);
    for (const node of childNodes) {
      if (node.nodeType === Node.TEXT_NODE && node.textContent && node.textContent.trim().length > 0) {
        const div = document.createElement('div');
        node.parentNode?.insertBefore(div, node);
        div.appendChild(node);
      }
    }

    // Clean up temporary dummy elements before recalculating
    pageEl.querySelectorAll('.page-break-dummy').forEach(el => el.remove());

    const children = Array.from(pageEl.children) as HTMLElement[];

    // Read all metrics in one pass to prevent layout thrashing (O(1) reflow instead of O(N))
    const metrics = children.map(child => ({
      child,
      top: child.offsetTop,
      height: child.offsetHeight,
      isBreak: child.style?.pageBreakAfter === 'always' ||
        child.classList?.contains('manual-page-break') ||
        (child.querySelector && child.querySelector('[style*="page-break-after: always"], .manual-page-break'))
    }));

    let currentPushOffset = 0;
    let forceNextPage = false;

    for (const m of metrics) {
      if (m.child.style.position === 'absolute') continue;
      
      let effectiveTop = m.top + currentPushOffset;
      let childBottom = effectiveTop + m.height;
      
      let currentPage = Math.floor(effectiveTop / 1080) + 1;
      const textTopLimit = (currentPage - 1) * 1080 + 96;
      const textBottomLimit = (currentPage - 1) * 1080 + 960;
      
      const tagName = m.child.tagName.toLowerCase();
      
      // If the element naturally lands inside the top margin of a page, push it down to the text start
      if (effectiveTop < textTopLimit) {
          const pushAmount = textTopLimit - effectiveTop;
          if (pushAmount > 0) {
              const dummy = document.createElement('div');
              dummy.className = 'page-break-dummy';
              dummy.contentEditable = 'false';
              dummy.style.height = pushAmount + 'px';
              dummy.style.display = 'flow-root';
              dummy.style.margin = '0';
              dummy.style.padding = '0';
              dummy.style.border = 'none';
              dummy.style.userSelect = 'none';
              m.child.parentNode?.insertBefore(dummy, m.child);
              currentPushOffset += pushAmount;
          }
          
          effectiveTop = textTopLimit;
          childBottom = effectiveTop + m.height;
      }
      
      if (tagName === 'table') {
         let tablePushOffset = 0;
         const rows = Array.from(m.child.querySelectorAll('tr'));
         const tableRectTop = m.child.getBoundingClientRect().top;
         
         for (const row of rows) {
            if (row.classList.contains('page-break-dummy')) continue;
            const rowRelTop = row.getBoundingClientRect().top - tableRectTop;
            let rowTop = effectiveTop + rowRelTop + tablePushOffset;
            let rowBottom = rowTop + row.offsetHeight;
            
            let rowPage = Math.floor(rowTop / 1080) + 1;
            let rowBottomLimit = (rowPage - 1) * 1080 + 960;
            
            if (rowBottom > rowBottomLimit) {
               rowPage = Math.floor(rowTop / 1080) + 1;
               const nextStart = rowPage * 1080 + 96;
               const push = nextStart - rowTop;
               
               if (push > 0) {
                   const dummy = document.createElement('tr');
                   dummy.className = 'page-break-dummy';
                   dummy.contentEditable = 'false';
                   dummy.style.userSelect = 'none';
                   dummy.innerHTML = `<td colspan="100" style="height: ${push}px; border: none; padding: 0;"></td>`;
                   row.parentNode?.insertBefore(dummy, row);
                   tablePushOffset += push;
               }
            }
         }
         
         m.height += tablePushOffset;
         childBottom += tablePushOffset;
      } else {
         if (forceNextPage || childBottom > textBottomLimit) {
            // Re-calculate currentPage in case the top margin push pushed it exactly to the border, though unlikely
            currentPage = Math.floor(effectiveTop / 1080) + 1;
            const nextPageTextStart = currentPage * 1080 + 96;
            const pushAmount = nextPageTextStart - effectiveTop;
            
            if (pushAmount > 0) {
                const dummy = document.createElement('div');
                dummy.className = 'page-break-dummy';
                dummy.contentEditable = 'false';
                dummy.style.height = pushAmount + 'px';
                dummy.style.display = 'flow-root'; // Prevent margin collapse
                dummy.style.margin = '0';
                dummy.style.padding = '0';
                dummy.style.border = 'none';
                dummy.style.userSelect = 'none';
                m.child.parentNode?.insertBefore(dummy, m.child);
                currentPushOffset += pushAmount;
            }
            
            forceNextPage = false;
         }
      }

      if (m.isBreak) {
        forceNextPage = true;
      }
    }
  }

  onPaste(e: ClipboardEvent) {
    // Attempt to sanitize HTML to remove problematic width/min-width inline styles
    // often found in pasted spreadsheet tables (Excel, Google Sheets).
    const htmlData = e.clipboardData?.getData('text/html');
    if (htmlData) {
      e.preventDefault();
      // Remove inline width and min-width attributes safely without consuming closing quotes
      let sanitizedHtml = htmlData.replace(/(?:min-|max-)?width\s*:\s*[^;"]+;?/ig, '')
                                  .replace(/(?:min-|max-)?height\s*:\s*[^;"]+;?/ig, '');
      
      // Also remove fixed width attributes from table elements
      sanitizedHtml = sanitizedHtml.replace(/ width="[^"]*"/ig, '')
                                   .replace(/ height="[^"]*"/ig, '');

      document.execCommand('insertHTML', false, sanitizedHtml);
    }
  }

  updatePageHeight() {
    if (!this.showPrintLayout) return;
    const el = document.querySelector('.page') as HTMLElement;
    if (el) {
      el.style.height = 'auto'; // Reset to measure natural content height
      const naturalHeight = el.scrollHeight;
      
      // Calculate exact number of pages needed 
      // Divide by 1080 because naturalHeight includes the 24px gaps injected by autoPaginate
      const pages = Math.max(1, Math.ceil(naturalHeight / 1080));
      
      // Snap container to exact boundaries including 24px gap between pages
      const exactHeight = (pages * 1056) + ((pages - 1) * 24);
      el.style.height = exactHeight + 'px';

      if (this.pageCountArray.length !== pages) {
        this.pageCountArray = new Array(pages).fill(0);
      }
    }
  }

  scrollToPage(index: number) {
    if (this.pageAreaRef && this.pageAreaRef.nativeElement) {
      // Each page is 1056px + 24px gap. 
      // The ruler and padding at the top of the scroll area might add offset, but scrolling exactly to page boundaries:
      const yOffset = index * (1056 + 24);
      this.pageAreaRef.nativeElement.scrollTo({ top: yOffset, behavior: 'smooth' });
    }
  }

  async deletePage(index: number, event: Event) {
    event.stopPropagation();
    if (!await this.showConfirm(`Are you sure you want to delete Page ${index + 1}?`)) return;

    const pageEl = document.querySelector('.page') as HTMLElement;
    if (!pageEl) return;

    const pageTopY = index * 1080;
    const pageBottomY = (index + 1) * 1080;

    const children = Array.from(pageEl.children) as HTMLElement[];
    let elementsDeleted = false;
    let lastChildOnPrevPage: HTMLElement | null = null;

    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      // If the child mathematically sits within this page's vertical coordinate block, remove it
      if (child.offsetTop >= pageTopY && child.offsetTop < pageBottomY) {
        if (!elementsDeleted && lastChildOnPrevPage) {
          const breakEls = lastChildOnPrevPage.classList.contains('manual-page-break')
            ? [lastChildOnPrevPage]
            : Array.from(lastChildOnPrevPage.querySelectorAll('.manual-page-break'));
          breakEls.forEach((el: any) => el.remove());
        }
        child.remove();
        elementsDeleted = true;
      } else if (child.offsetTop < pageTopY) {
        lastChildOnPrevPage = child;
      }
    }

    if (!elementsDeleted && lastChildOnPrevPage) {
      const breakEls = (lastChildOnPrevPage as any).classList.contains('manual-page-break')
        ? [lastChildOnPrevPage]
        : Array.from((lastChildOnPrevPage as any).querySelectorAll('.manual-page-break'));
      if (breakEls.length > 0) {
        breakEls.forEach((el: any) => (el as HTMLElement).remove());
        elementsDeleted = true;
      }
    }

    if (elementsDeleted) {
      this.save(); // Save will trigger autoPaginate() which recalculates all margins instantly
      this.updatePageHeight(); // Shrink the physical container and update sidebar thumbnails
      this.showToast(`Page ${index + 1} deleted.`);
    } else {
      this.updatePageHeight(); // Clean up any ghost pages
      this.showToast('Page is already empty.');
    }
  }

  @HostListener('window:resize') 
  onWindowResize() { this.updateOverlay(); }

  toggleTextBoxSidebar(show: boolean) {
    this.showTextBoxOptions = show;
    setTimeout(() => this.updateOverlay(), 50);
  }

  toggleAutomateSidebar(type: string) {
    if (!type) {
      this.showAutomateSidebar = false;
      return;
    }
    
    this.closeMenus();
    this.showAutomateSidebar = true;
    this.automateTitle = type;
    
    if (type === 'Merge Template') {
       this.automateSubtitle = 'Generate documents in bulk';
       this.automateDescription = 'Connect a data source such as a CSV file or Vmail Sheet to automatically populate merge fields in this document.';
    } else if (type === 'Fillable Template') {
       this.automateSubtitle = 'Create interactive forms';
       this.automateDescription = 'Drag and drop fillable form fields onto the document canvas to collect information from recipients.';
    } else if (type === 'Sign Template') {
       this.automateSubtitle = 'Collect digital signatures';
       this.automateDescription = 'Add signer roles, drag signature fields into the document, and set up automated delivery workflows.';
    }
  }

  @HostListener('document:dblclick', ['$event'])
  onDoubleClick(e: MouseEvent) {
    const target = e.target as HTMLElement;
    if (target.classList.contains('vmail-text-box') || target.closest('.vmail-text-box')) {
      this.toggleTextBoxSidebar(true);
    }
  }

  @HostListener('document:keydown', ['$event'])
  async onKeyDown(e: KeyboardEvent) {
    if (e.key === ' ' && this.activeAutocorrect) {
      const sel = window.getSelection();
      if (sel && sel.focusNode && sel.focusNode.nodeType === 3) {
         const focusNode = sel.focusNode;
         let el: any = focusNode.parentNode;
         let inEditor = false;
         while (el) {
           if (el.classList && el.classList.contains('page')) { inEditor = true; break; }
           el = el.parentNode;
         }
         if (inEditor) {
            const text = focusNode.nodeValue || '';
            const offset = sel.focusOffset;
            let start = offset;
            while (start > 0 && /\S/.test(text[start - 1])) start--;
            const word = text.substring(start, offset);
            
            const corrections: any = {
              'teh': 'the',
              'hte': 'the',
              'te': 'the',
              'dont': "don't",
              'cant': "can't",
              'im': "I'm",
              'thats': "that's",
              'thn': 'then',
              'taht': 'that',
              'tht': 'that',
              'watn': 'want',
              'jsut': 'just',
              'woudl': 'would',
              'becuase': 'because',
              'definately': 'definitely',
              'seperate': 'separate',
              'occured': 'occurred',
              'alot': 'a lot',
              'recieve': 'receive',
              'untill': 'until',
              'ur': 'your',
              '(c)': '©',
              '(r)': '®',
              '(tm)': '™',
              '->': '→',
              '<-': '←'
            };
            
            if (corrections[word.toLowerCase()]) {
               e.preventDefault();
               const corrected = corrections[word.toLowerCase()];
               const isCapitalized = word[0] === word[0].toUpperCase();
               const finalWord = isCapitalized ? corrected.charAt(0).toUpperCase() + corrected.slice(1) : corrected;
               
               const range = document.createRange();
               range.setStart(focusNode, start);
               range.setEnd(focusNode, offset);
               range.deleteContents();
               
               const textNode = document.createTextNode(finalWord + ' ');
               range.insertNode(textNode);
               
               const newRange = document.createRange();
               newRange.setStartAfter(textNode);
               newRange.collapse(true);
               sel.removeAllRanges();
               sel.addRange(newRange);
               return;
            } else if (word.length >= 3) {
               fetch('https://api.languagetool.org/v2/check', {
                 method: 'POST',
                 headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                 body: `language=en-US&text=${encodeURIComponent(word)}`
               })
               .then(r => r.json())
               .then(data => {
                  if (data.matches && data.matches.length > 0 && data.matches[0].replacements && data.matches[0].replacements.length > 0) {
                     if (data.matches[0].rule && data.matches[0].rule.issueType === 'misspelling') {
                         const corrected = data.matches[0].replacements[0].value;
                         if (corrected.toLowerCase() === word.toLowerCase()) return;
                         
                         const isCapitalized = word[0] === word[0].toUpperCase();
                         const finalWord = isCapitalized ? corrected.charAt(0).toUpperCase() + corrected.slice(1) : corrected;
                         
                         if (focusNode.nodeValue) {
                            const val = focusNode.nodeValue;
                            if (val.substring(start, offset) === word) {
                                const s = window.getSelection();
                                let savedRange = s && s.rangeCount > 0 ? s.getRangeAt(0).cloneRange() : null;
                                const diff = finalWord.length - word.length;
                                
                                focusNode.nodeValue = val.substring(0, start) + finalWord + val.substring(offset);
                                
                                if (savedRange && savedRange.startContainer === focusNode) {
                                    try {
                                        const newStart = savedRange.startOffset > offset ? savedRange.startOffset + diff : savedRange.startOffset;
                                        const newEnd = savedRange.endOffset > offset ? savedRange.endOffset + diff : savedRange.endOffset;
                                        const r = document.createRange();
                                        r.setStart(focusNode, newStart);
                                        r.setEnd(focusNode, newEnd);
                                        s?.removeAllRanges();
                                        s?.addRange(r);
                                    } catch(e) {}
                                }
                            }
                         }
                     }
                  }
               }).catch(err => console.log('Autocorrect API error', err));
            }
         }
      }
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      this.save(true);
      return;
    }



    if (this.selectedObject && (e.key === 'Backspace' || e.key === 'Delete')) {
      if (this.selectedObject.tagName === 'IMG') {
        this.selectedObject.remove();
        this.selectedObject = null;
        this.updateOverlay();
        this.save();
        e.preventDefault();
        e.stopPropagation();
      } else if (this.selectedObject.classList.contains('vmail-text-box')) {
        if (e.key === 'Delete') {
          const sel = window.getSelection();
          if (sel && sel.focusNode === this.selectedObject.parentNode) {
            this.selectedObject.remove();
            this.selectedObject = null;
            this.toggleTextBoxSidebar(false);
            this.updateOverlay();
            this.save();
            e.preventDefault();
            e.stopPropagation();
          }
        }
      }
    }

    if (e.key === 'Backspace') {
      const target = e.target as HTMLElement;
      const textBox = target.classList?.contains('vmail-text-box') ? target : target.closest('.vmail-text-box');
      if (textBox) {
        if (textBox.innerHTML === '' || textBox.innerHTML === '<br>' || textBox.childNodes.length === 0) {
          e.preventDefault(); // Prevent native browser from deleting the empty text box div
        }
      }
    }
  }

  onPageMouseDown(e: MouseEvent) {
    const target = e.target as HTMLElement;
    
    // If we click a resize handle, do not reset the selected image
    if (target.classList.contains('resize-handle')) {
      return; 
    }

    // Handle clicks on empty space or unselectable dummy spacers
    if (target.classList.contains('page-break-dummy') || target.classList.contains('page') || target.classList.contains('page-area')) {
      // Allow default but ensure editor is focused
      setTimeout(() => {
        if (this.editor && this.editor.nativeElement) {
          const editor = this.editor.nativeElement;
          editor.focus();
          
          // If clicking a dummy, try to place cursor at the end of the preceding paragraph
          if (target.classList.contains('page-break-dummy') && target.previousSibling) {
             const sel = window.getSelection();
             const range = document.createRange();
             range.selectNodeContents(target.previousSibling as Node);
             range.collapse(false); // false means end of node
             sel?.removeAllRanges();
             sel?.addRange(range);
          } else if (target.classList.contains('page') || target.classList.contains('page-area')) {
             // Fallback: place cursor at very end of document if no valid selection exists
             const sel = window.getSelection();
             if (!sel || !sel.focusNode || sel.focusNode === editor || sel.focusNode.nodeName === 'BODY') {
                const range = document.createRange();
                range.selectNodeContents(editor);
                range.collapse(false);
                sel?.removeAllRanges();
                sel?.addRange(range);
             }
          }
        }
      }, 10);
    }

    if (target.tagName === 'IMG' || target.classList.contains('vmail-text-box') || target.closest('.vmail-text-box')) {
      const actualTarget = target.classList.contains('vmail-text-box') ? target : target.closest('.vmail-text-box') as HTMLElement || target;
      e.preventDefault(); // Prevent native drag-and-drop
      
      this.selectedObject = actualTarget as HTMLElement;
      
      // Do not select the image text node, so pressing Space does not delete it
      if (this.selectedObject.tagName !== 'IMG') {
        const sel = window.getSelection();
        if (sel) {
          const range = document.createRange();
          range.selectNode(actualTarget);
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }

      // Initiate custom free drag only if it is already absolute
      if (this.selectedObject.style.position === 'absolute') {
        this.isDraggingImage = true;
        this.dragStartX = e.clientX;
        this.dragStartY = e.clientY;
        this.imgStartX = parseFloat(this.selectedObject.style.left || '0');
        this.imgStartY = parseFloat(this.selectedObject.style.top || '0');
      }

      this.nativeDragStartWidth = this.selectedObject.offsetWidth;
      this.nativeDragStartHeight = this.selectedObject.offsetHeight;
      
      this.updateOverlay();
    } else {
      if (this.selectedObject) {
        this.selectedObject = null;
        this.updateOverlay();
      }
    }
  }

  updateOverlay() {
    if (!this.selectedObject) return;
    const rect = this.selectedObject.getBoundingClientRect();
    this.overlayRect = {
      top: rect.top,
      left: rect.left,
      width: rect.width,
      height: rect.height
    };
  }

  startResize(e: MouseEvent, corner: string) {
    e.preventDefault();
    e.stopPropagation();
    this.isResizing = true;
    this.resizeCorner = corner;
    this.startX = e.clientX;
    this.startY = e.clientY;
    this.startW = this.selectedObject!.offsetWidth;
    this.startH = this.selectedObject!.offsetHeight;
    this.imgStartX = parseFloat(this.selectedObject!.style.left || '0');
    this.imgStartY = parseFloat(this.selectedObject!.style.top || '0');
  }

  @HostListener('document:mousemove', ['$event'])
  onMouseMove(e: MouseEvent) {
    if (this.isResizing && this.selectedObject) {
      const dx = e.clientX - this.startX;
      let newW = this.startW;
      
      const dy = e.clientY - this.startY;
      let newH = this.startH;
      let newLeft = this.imgStartX;
      let newTop = this.imgStartY;

      if (this.resizeCorner === 'br') {
        newW = Math.max(10, this.startW + dx);
        newH = Math.max(10, this.startH + dy);
      } else if (this.resizeCorner === 'tr') {
        newW = Math.max(10, this.startW + dx);
        newH = Math.max(10, this.startH - dy);
        newTop = this.imgStartY + (this.startH - newH);
      } else if (this.resizeCorner === 'bl') {
        newW = Math.max(10, this.startW - dx);
        newH = Math.max(10, this.startH + dy);
        newLeft = this.imgStartX + (this.startW - newW);
      } else if (this.resizeCorner === 'tl') {
        newW = Math.max(10, this.startW - dx);
        newH = Math.max(10, this.startH - dy);
        newLeft = this.imgStartX + (this.startW - newW);
        newTop = this.imgStartY + (this.startH - newH);
      }

      this.selectedObject.style.width = newW + 'px';
      
      if (this.selectedObject.tagName === 'IMG') {
        this.selectedObject.style.height = 'auto'; 
      } else {
        this.selectedObject.style.height = newH + 'px';
      }

      if (this.selectedObject.style.position === 'absolute') {
        this.selectedObject.style.left = newLeft + 'px';
        this.selectedObject.style.top = newTop + 'px';
      }

      this.updateOverlay();
    } else if (this.isDraggingImage && this.selectedObject) {
      let newLeft = this.imgStartX + (e.clientX - this.dragStartX);
      let newTop = this.imgStartY + (e.clientY - this.dragStartY);
      
      // Constrain to physical page boundaries
      const pageEl = document.querySelector('.page') as HTMLElement;
      if (pageEl) {
        const maxLeft = pageEl.offsetWidth - this.selectedObject.offsetWidth;
        const maxTop = pageEl.offsetHeight - this.selectedObject.offsetHeight;
        newLeft = Math.max(0, Math.min(newLeft, maxLeft));
        newTop = Math.max(0, Math.min(newTop, maxTop));
      }

      this.selectedObject.style.left = newLeft + 'px';
      this.selectedObject.style.top = newTop + 'px';
      this.updateOverlay();
    } else if (this.cropModalVisible && this.cropAction && this.cropArea && this.cropStartArea && this.cropImageEl) {
      const dx = e.clientX - this.cropStartX;
      const dy = e.clientY - this.cropStartY;
      const img = this.cropImageEl.nativeElement;

      if (this.cropAction === 'drag') {
         let newX = this.cropStartArea.x + dx;
         let newY = this.cropStartArea.y + dy;
         newX = Math.max(img.offsetLeft, Math.min(newX, img.offsetLeft + img.width - this.cropArea.w));
         newY = Math.max(img.offsetTop, Math.min(newY, img.offsetTop + img.height - this.cropArea.h));
         this.cropArea.x = newX;
         this.cropArea.y = newY;
      } else {
         let newX = this.cropStartArea.x;
         let newY = this.cropStartArea.y;
         let newW = this.cropStartArea.w;
         let newH = this.cropStartArea.h;
         
         if (this.cropAction === 'br') {
           newW = Math.max(20, this.cropStartArea.w + dx);
           newH = Math.max(20, this.cropStartArea.h + dy);
         } else if (this.cropAction === 'tr') {
           newW = Math.max(20, this.cropStartArea.w + dx);
           newH = Math.max(20, this.cropStartArea.h - dy);
           newY = this.cropStartArea.y + (this.cropStartArea.h - newH);
         } else if (this.cropAction === 'bl') {
           newW = Math.max(20, this.cropStartArea.w - dx);
           newH = Math.max(20, this.cropStartArea.h + dy);
           newX = this.cropStartArea.x + (this.cropStartArea.w - newW);
         } else if (this.cropAction === 'tl') {
           newW = Math.max(20, this.cropStartArea.w - dx);
           newH = Math.max(20, this.cropStartArea.h - dy);
           newX = this.cropStartArea.x + (this.cropStartArea.w - newW);
           newY = this.cropStartArea.y + (this.cropStartArea.h - newH);
         }

         if (newX < img.offsetLeft) { newW -= (img.offsetLeft - newX); newX = img.offsetLeft; }
         if (newY < img.offsetTop) { newH -= (img.offsetTop - newY); newY = img.offsetTop; }
         if (newX + newW > img.offsetLeft + img.width) newW = img.offsetLeft + img.width - newX;
         if (newY + newH > img.offsetTop + img.height) newH = img.offsetTop + img.height - newY;

         this.cropArea.x = newX;
         this.cropArea.y = newY;
         this.cropArea.w = newW;
         this.cropArea.h = newH;
      }
    }
  }

  @HostListener('document:mouseup', ['$event'])
  onMouseUp(e: MouseEvent) {
    this.cropAction = null;
    if (this.isResizing || this.isDraggingImage) {
      const wasDragging = this.isDraggingImage;
      this.isResizing = false;
      this.isDraggingImage = false;
      
      if (wasDragging && this.selectedObject && this.selectedObject.tagName === 'IMG') {
          const img = this.selectedObject;
          
          if (img.dataset['wrapStyle'] === 'absolute') {
              // Leave it exactly where it is (it's already position: absolute)
          } else {
              // Disable hit testing on the moving image so the browser can find the text underneath
              const oldPointerEvents = img.style.pointerEvents;
              img.style.pointerEvents = 'none';
              
              let targetNode: Node | null = null;
              let targetOffset = 0;
              
              if ((document as any).caretPositionFromPoint) {
                  const pos = (document as any).caretPositionFromPoint(e.clientX, e.clientY);
                  if (pos) { targetNode = pos.offsetNode; targetOffset = pos.offset; }
              } else if (document.caretRangeFromPoint) {
                  const range = document.caretRangeFromPoint(e.clientX, e.clientY);
                  if (range) { targetNode = range.startContainer; targetOffset = range.startOffset; }
              }
              
              // Restore image styles
              img.style.pointerEvents = oldPointerEvents;
              img.style.position = '';
              img.style.left = '';
              img.style.top = '';
              
              const pageEl = document.querySelector('.page') as HTMLElement;
              if (pageEl) {
                  let isSafe = true;
                  let curr = targetNode;
                  while(curr && curr !== pageEl) {
                      if (curr.nodeType === Node.ELEMENT_NODE && (curr as HTMLElement).classList.contains('page-break-dummy')) {
                          isSafe = false;
                          break;
                      }
                      curr = curr.parentNode;
                  }
                  
                  if (targetNode && pageEl.contains(targetNode) && isSafe) {
                      if (targetNode.nodeType === Node.TEXT_NODE) {
                          const split = (targetNode as Text).splitText(targetOffset);
                          targetNode.parentNode?.insertBefore(img, split);
                      } else {
                          if (targetNode.childNodes.length > targetOffset) {
                              targetNode.insertBefore(img, targetNode.childNodes[targetOffset]);
                          } else {
                              targetNode.appendChild(img);
                          }
                      }
                  }
              }
          }
          this.updateOverlay();
      }

      this.save();
    } else if (this.selectedObject && this.selectedObject.tagName === 'IMG') {
      // Catch native image resizing that doesn't trigger isResizing
      if (this.selectedObject.offsetWidth !== this.nativeDragStartWidth || this.selectedObject.offsetHeight !== this.nativeDragStartHeight) {
          setTimeout(() => {
              this.autoPaginate();
              this.updatePageHeight();
              this.save();
          }, 100);
      }
    }
  }

  setWrapStyle(style: string) {
    if (this.selectedObject && this.selectedObject.tagName === 'IMG') {
       this.selectedObject.dataset['wrapStyle'] = style;
       
       if (style === 'inline') {
         this.selectedObject.style.position = '';
         this.selectedObject.style.float = 'none';
         this.selectedObject.style.display = 'inline-block';
         this.selectedObject.style.margin = '0';
       } else if (style === 'left') {
         this.selectedObject.style.position = '';
         this.selectedObject.style.float = 'left';
         this.selectedObject.style.display = 'inline-block';
         this.selectedObject.style.margin = '0 16px 8px 0';
       } else if (style === 'right') {
         this.selectedObject.style.position = '';
         this.selectedObject.style.float = 'right';
         this.selectedObject.style.display = 'inline-block';
         this.selectedObject.style.margin = '0 0 8px 16px';
       } else if (style === 'break') {
         this.selectedObject.style.position = '';
         this.selectedObject.style.float = 'none';
         this.selectedObject.style.display = 'block';
         this.selectedObject.style.margin = '16px auto';
       } else if (style === 'absolute') {
         // Switch to absolute positioning
         const pageEl = document.querySelector('.page') as HTMLElement;
         if (this.selectedObject.style.position !== 'absolute' && pageEl) {
           const pageRect = pageEl.getBoundingClientRect();
           const imgRect = this.selectedObject.getBoundingClientRect();
           this.selectedObject.style.position = 'absolute';
           this.selectedObject.style.left = (imgRect.left - pageRect.left) + 'px';
           this.selectedObject.style.top = (imgRect.top - pageRect.top) + 'px';
           this.selectedObject.style.margin = '0';
           this.selectedObject.style.float = 'none';
         }
       }
       this.updateOverlay();
       this.save();
    }
  }

  openCropModal() {
    if (this.selectedObject && this.selectedObject.tagName === 'IMG') {
      this.cropTargetImg = this.selectedObject as HTMLImageElement;
      this.cropTargetSrc = this.cropTargetImg.src;
      this.cropModalVisible = true;
    }
  }

  initCropArea() {
    if (!this.cropImageEl) return;
    const img = this.cropImageEl.nativeElement;
    this.cropArea = {
      x: img.offsetLeft + img.width * 0.1,
      y: img.offsetTop + img.height * 0.1,
      w: img.width * 0.8,
      h: img.height * 0.8
    };
  }

  startCropDrag(e: MouseEvent) {
    e.preventDefault();
    if (!this.cropArea) return;
    this.cropAction = 'drag';
    this.cropStartX = e.clientX;
    this.cropStartY = e.clientY;
    this.cropStartArea = { ...this.cropArea };
  }

  startCropResize(e: MouseEvent, corner: string) {
    e.preventDefault();
    e.stopPropagation();
    if (!this.cropArea) return;
    this.cropAction = corner;
    this.cropStartX = e.clientX;
    this.cropStartY = e.clientY;
    this.cropStartArea = { ...this.cropArea };
  }

  applyCrop() {
    if (!this.cropTargetImg || !this.cropArea || !this.cropImageEl) return;
    const img = this.cropImageEl.nativeElement;
    
    const scaleX = img.naturalWidth / img.width;
    const scaleY = img.naturalHeight / img.height;

    const sourceX = (this.cropArea.x - img.offsetLeft) * scaleX;
    const sourceY = (this.cropArea.y - img.offsetTop) * scaleY;
    const sourceW = this.cropArea.w * scaleX;
    const sourceH = this.cropArea.h * scaleY;

    const canvas = document.createElement('canvas');
    canvas.width = sourceW;
    canvas.height = sourceH;
    const ctx = canvas.getContext('2d');
    if (ctx) {
       ctx.drawImage(img, sourceX, sourceY, sourceW, sourceH, 0, 0, sourceW, sourceH);
       this.cropTargetImg.src = canvas.toDataURL('image/png');
       
       this.cropTargetImg.style.width = this.cropArea.w + 'px';
       this.cropTargetImg.style.height = this.cropArea.h + 'px';
       
       this.updateOverlay();
       this.save();
    }
    this.cropModalVisible = false;
  }

  onAutoSaveChange() {
    if (this.autoSaveEnabled) {
      this.save(true);
    }
  }

  toggleAutoSave() {
    this.autoSaveEnabled = !this.autoSaveEnabled;
    this.onAutoSaveChange();
  }

  save(force: boolean = false) {
    this.autoPaginate();
    const el = document.querySelector('.page') as HTMLElement;
    if (el) {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = el.innerHTML;
      tempDiv.querySelectorAll('.page-break-dummy').forEach(row => row.remove());
      
      this.htmlContent = tempDiv.innerHTML;
      if (!this.htmlContent.trim() || this.htmlContent.trim() === '<div><br></div>') {
         this.htmlContent = '<div><br></div>';
         el.innerHTML = this.htmlContent;
      }
    }
    
    // Always broadcast via websocket so memory state stays in sync across clients,
    // passing the appropriate autosave flag so the backend knows whether to mark it dirty.
    const shouldSaveToDisk = force || this.autoSaveEnabled;
    this.api.sendUpdate(JSON.stringify({ html: this.htmlContent }), this.title, shouldSaveToDisk);
    
    if (!shouldSaveToDisk) {
      return;
    }
    
    this.isSaving = true;
    this.api.saveDocument(this.docId, this.title, JSON.stringify({ html: this.htmlContent })).subscribe(() => {
      this.isSaving = false;
      this.lastSavedTime = new Date();
    });
  }

  onShareSearch() {
    if (this.shareQuery.trim().length < 2) {
      this.userSearchResults = [];
      return;
    }
    this.api.searchUsers(this.shareQuery).subscribe(users => {
      this.userSearchResults = users.filter((u: any) => u.id !== this.auth.user?.id);
    });
  }

  selectShareUser(user: any) {
    this.shareQuery = user.email;
    this.userSearchResults = [];
  }

  performShare() {
    if (!this.shareQuery) return;
    this.api.shareDocument(this.docId, this.shareQuery, this.shareRole.toLowerCase()).subscribe({
      next: () => {
        this.showToast(`Shared with ${this.shareQuery}`);
        this.shareQuery = '';
    this.shareModalOpen = false;
      },
      error: () => this.showToast('Failed to share: User not found.')
    });
  }

  makePublic() {
    this.isPublic = true;
    this.copyLink();
  }

  copyLink() {
    this.closeMenus();
    this.shareModalOpen = false;
    navigator.clipboard.writeText(window.location.href)
      .then(() => this.showToast('Link copied! Anyone with the link can collaborate.'));
  }

  triggerImport(event?: Event) {
    if (event) {
      event.stopPropagation();
    }
    if (this.importInput) {
      this.importInput.nativeElement.click();
    }
    this.closeMenus();
  }

  connectCloud(provider: string) {
    this.isConnectingCloud = true;
    this.connectingCloudName = provider;
    setTimeout(() => {
      this.isConnectingCloud = false;
      this.cloudModalOpen = false;
      // Trigger local file import as a functional mock for cloud import
      this.triggerImport();
    }, 1500);
  }

  importFile(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.showToast(`Importing ${file.name}...`);
      this.api.importFile(file, this.docId).subscribe({
        next: (doc: any) => {
          console.log('[IMPORT] Response received:', {
            id: doc.id,
            title: doc.title,
            contentLength: doc.content?.length,
            contentPreview: doc.content?.substring(0, 200)
          });
          try {
            let p = JSON.parse(doc.content || '{}');
            console.log('[IMPORT] Parsed content keys:', Object.keys(p));
            if (Array.isArray(p) && p.length > 0) p = p[0];

            if (p.html) {
              console.log('[IMPORT] HTML found, length:', p.html.length);
              this.htmlContent = this.sanitizeImportedHtml(p.html);

              // ── SAVE directly with the parsed content string (not from DOM) ──
              // Saving from DOM (save()) is unreliable on live because the browser
              // may not have rendered el.innerHTML yet when save() fires, causing
              // blank content to overwrite the real import data in R2 storage.
              this.title = doc.title;
              this.isSaving = true;
              this.api.saveDocument(doc.id, doc.title, JSON.stringify({ html: this.htmlContent })).subscribe({
                next: () => {
                  this.isSaving = false;
                  this.lastSavedTime = new Date();
                  console.log('[IMPORT] Content saved to R2 successfully.');
                },
                error: (err: any) => {
                  this.isSaving = false;
                  console.error('[IMPORT] Failed to save content to R2:', err);
                }
              });
              this.api.sendUpdate(JSON.stringify({ html: this.htmlContent }), doc.title, true);

              // ── Apply to DOM and paginate ──
              const el = document.querySelector('.page') as HTMLElement;
              if (el) {
                el.innerHTML = this.htmlContent;
                console.log('[IMPORT] DOM updated, innerHTML length:', el.innerHTML.length);

                // Wait for images to load before paginating
                const images = Array.from(el.querySelectorAll('img'));
                let loadedCount = 0;
                const checkFinished = () => {
                  loadedCount++;
                  if (loadedCount >= images.length) {
                    setTimeout(() => this.autoPaginate(), 150);
                  }
                };
                if (images.length === 0) {
                  setTimeout(() => this.autoPaginate(), 150);
                } else {
                  images.forEach(img => {
                    if (img.complete) checkFinished();
                    else { img.onload = checkFinished; img.onerror = checkFinished; }
                  });
                }
              }
            } else {
              console.warn('[IMPORT] No html field in parsed content. Keys:', Object.keys(p));
              this.showToast('Import complete but no content found. Check console for details.');
            }
          } catch (e) {
            console.error('[IMPORT] Error parsing document content:', e, 'Raw:', doc.content?.substring(0, 500));
            this.showToast('Import parsing error. Check console.');
          }
          this.title = doc.title;
          this.showToast(`${file.name} imported successfully.`);
        },
        error: (err: any) => {
          this.showToast(`Error importing ${file.name}.`);
          console.error('[IMPORT] HTTP error:', err);
        }
      });
      event.target.value = '';
    }
  }

  saveAs() {
    this.closeMenus();
    this.saveAsType = 'document';
    this.saveAsInput = this.title + ' Copy';
    this.saveAsModalOpen = true;
  }

  saveAsTemplate() {
    this.closeMenus();
    this.saveAsType = 'template';
    this.saveAsInput = this.title + ' (Template)';
    this.saveAsModalOpen = true;
  }

  confirmSaveAs() {
    const newTitle = this.saveAsInput.trim();
    if (!newTitle) {
      this.showToast('Please enter a name');
      return;
    }
    this.saveAsModalOpen = false;
    this.api.createDocument(newTitle, 'writer').subscribe((res: any) => {
      this.api.saveDocument(res.id, newTitle, JSON.stringify({ html: this.htmlContent })).subscribe(() => {
        if (this.saveAsType === 'template') {
          this.showToast(`Saved as Template: ${newTitle}`);
        } else {
          window.open(`/doc/${res.id}`, '_blank');
          this.showToast(`Saved as ${newTitle}`);
        }
      });
    });
  }

  openDoc(id: string) {
    this.closeMenus();
    window.open(`/doc/${id}`, '_blank');
  }

  exportFile(format: string, password?: string) {
    this.save();
    this.api.exportDocument(this.docId, format, password).subscribe(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `${this.title}.${format}`; a.click();
      URL.revokeObjectURL(url);
    });
  }

  openPasswordModal(format: string) {
    this.closeMenus();
    this.passwordFormat = format;
    this.passwordInput = '';
    this.passwordModalOpen = true;
  }

  submitPassword() {
    if (!this.passwordInput.trim()) {
      this.showToast('Please enter a password');
      return;
    }
    this.passwordModalOpen = false;
    this.exportFile(this.passwordFormat, this.passwordInput.trim());
  }

  checkSpelling() {
    this.closeMenus();
    this.showToast('Browser spellcheck is active. Right-click underlined words for suggestions.');
  }

  openDictionary() {
    this.closeMenus();
    const selection = window.getSelection();
    this.dictWord = selection ? selection.toString().trim() : '';
    this.dictModalOpen = true;
    this.dictResults = [];
    this.dictError = '';
    if (this.dictWord) {
      this.searchDictionary();
    }
  }



  searchDictionary() {
    if (!this.dictWord.trim()) return;
    this.dictLoading = true;
    this.dictError = '';
    this.dictResults = [];
    
    fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(this.dictWord.trim())}`)
      .then(res => {
        if (!res.ok) throw new Error('Word not found');
        return res.json();
      })
      .then(data => {
        this.dictResults = data;
        this.dictLoading = false;
      })
      .catch(err => {
        this.dictError = 'Definition not found.';
        this.dictLoading = false;
      });
  }

  showToast(msg: string) {
    this.toastMsg = msg; this.toastVisible = true;
    setTimeout(() => this.toastVisible = false, 2500);
  }

  back() { this.save(); this.router.navigate(['/']); }

  zoomIn() {
    if (this.zoomLevel < 200) {
      this.zoomLevel += 10;
      const page = document.querySelector('.page') as HTMLElement;
      if (page) page.style.transform = `scale(${this.zoomLevel / 100})`;
    }
  }

  zoomOut() {
    if (this.zoomLevel > 50) {
      this.zoomLevel -= 10;
      const page = document.querySelector('.page') as HTMLElement;
      if (page) page.style.transform = `scale(${this.zoomLevel / 100})`;
    }
  }

  updateCounts() {
    const el = document.querySelector('.page') as HTMLElement;
    if (el) {
      const text = el.innerText || '';
      this.charCount = text.length;
      this.wordCount = text.trim() ? text.trim().split(/\s+/).length : 0;
    }
  }

  ngOnDestroy() {
    document.removeEventListener('selectionchange', this.onSelectionChange.bind(this));
    document.removeEventListener('click', this.globalClickHandler);
    this.syncSub?.unsubscribe();
    this.api.disconnectSync();
  }
}

