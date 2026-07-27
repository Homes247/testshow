import { Component, Input, Output, EventEmitter, OnInit, OnDestroy, OnChanges, SimpleChanges, ViewChild, ElementRef, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ApiService } from '../../services/api.service';
import { AuthService } from '../../services/auth.service';
import { ChatSocketService } from '../../services/chat-socket.service';
import { HttpClient } from '@angular/common/http';
import { environment } from '../../../environments/environment';
import { Subscription } from 'rxjs';
import { MediaPickerComponent } from '../media-picker/media-picker.component';

@Component({
  selector: 'app-chat-widget',
  standalone: true,
  imports: [CommonModule, FormsModule, MediaPickerComponent],
  template: `
    <div class="widget-panel shadow-lg" *ngIf="activeWidget" (click)="onPanelClick($event)">
      
      <!-- CHAT VIEW -->
      <ng-container *ngIf="activeWidget === 'chat' && !openedChatId">
        <div class="wp-header-chat">
          <div style="display:flex; align-items:center; gap:12px;">
            <div class="chat-av">
                <img *ngIf="currentUser?.avatar_url" [src]="currentUser.avatar_url" onerror="this.style.display='none'">
                <span *ngIf="!currentUser?.avatar_url" style="color:#10b981;">ME</span>
            </div>
            <div>
              <div style="font-size:14px; font-weight:600; color:#fff;">{{ currentUser?.name || 'admin' }}</div>
              <div style="font-size:11px; color:#d2e3fc; display:flex; align-items:center; gap:4px;">
                <div style="width:6px; height:6px; background:#4ade80; border-radius:50%;"></div> Online
              </div>
            </div>
          </div>
          <span class="material-symbols-outlined chat-close" (click)="close.emit()">close</span>
        </div>
        <div class="wp-search">
          <span class="material-symbols-outlined">search</span>
          <input type="text" placeholder="Search chats or contacts..." [(ngModel)]="searchQuery" (ngModelChange)="onSearchChange($event)">
        </div>
        <div class="wp-section-title" *ngIf="filteredConversations().length > 0">RECENT CHATS</div>
        <div class="wp-body-list" *ngIf="filteredConversations().length > 0">
          <div class="chat-list-item" *ngFor="let c of filteredConversations()" (click)="openChat(c.other_user.id)">
            <div class="cli-av" [style.background]="isSelf(c.other_user.id) ? '#10b981' : (c.other_user.avatar_color || '#6366f1')">
              <img *ngIf="c.other_user.avatar_url && !isSelf(c.other_user.id)" [src]="c.other_user.avatar_url">
              <span *ngIf="!c.other_user.avatar_url && !isSelf(c.other_user.id)">{{ getInitials(c.other_user.name) }}</span>
              <span *ngIf="isSelf(c.other_user.id)" style="font-size:12px;">ME</span>
              <div [class]="'cli-status ' + (isOnline(c.other_user.last_login) ? 'online' : 'offline')"></div>
            </div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ c.other_user.name }}</span>
                <span class="cli-time">{{ formatTime(c.last_time) }}</span>
              </div>
              <div class="cli-preview" [style.color]="isSelf(c.other_user.id) ? '#10b981' : 'var(--text-secondary)'" [style.font-style]="isSelf(c.other_user.id) ? 'italic' : 'normal'">
                <span [innerHTML]="isSelf(c.other_user.id) ? 'Personal Space' : linkify(cleanMessage(c.last_message || 'No messages yet'))"></span>
              </div>
            </div>
            <div *ngIf="c.unread > 0" style="background:#ef4444; color:white; font-size:10px; font-weight:bold; border-radius:10px; padding:2px 6px; margin-left:8px;">{{c.unread}}</div>
          </div>
          <div *ngIf="conversations.length === 0" style="padding: 20px; text-align:center; color:var(--text-secondary); font-size:13px;">
            No recent chats
          </div>
        </div>
        
        <div class="wp-section-title" *ngIf="searchQuery && filteredContacts().length > 0" style="margin-top:8px;">ALL CONTACTS</div>
        <div class="wp-body-list" *ngIf="searchQuery && filteredContacts().length > 0">
          <div class="chat-list-item" *ngFor="let u of filteredContacts()" (click)="openWidgetChat(u.id)">
            <div class="cli-av" [style.background]="u.avatar_color || '#6366f1'">
              <img *ngIf="u.avatar_url" [src]="u.avatar_url">
              <span *ngIf="!u.avatar_url">{{ getInitials(u.name) }}</span>
              <div [class]="'cli-status ' + (isOnline(u.last_login) ? 'online' : 'offline')"></div>
            </div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ isSelf(u.id) ? 'Saved Messages' : u.name }}</span>
              </div>
              <div class="cli-preview" style="color:var(--text-secondary)">{{ isSelf(u.id) ? 'Personal Space' : 'Team Member' }}</div>
            </div>
          </div>
          <div class="chat-list-item" *ngFor="let u of serverContacts" (click)="openWidgetChat(u.id)">
            <div class="cli-av" [style.background]="u.avatar_color || '#6366f1'">
              <img *ngIf="u.avatar_url" [src]="u.avatar_url">
              <span *ngIf="!u.avatar_url">{{ getInitials(u.name) }}</span>
            </div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ u.name }}</span>
              </div>
              <div class="cli-preview" style="color:var(--text-secondary)">{{ u.email }}</div>
            </div>
          </div>
        </div>
      </ng-container>

      <!-- OPENED CHAT CONVERSATION VIEW -->
      <ng-container *ngIf="activeWidget === 'chat' && openedChatId">
        <div class="wp-header-chat">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined icon-btn" (click)="closeChat()">arrow_back</span>
            <div class="cli-av" style="width:32px; height:32px; font-size:12px; margin-right:4px;" [style.background]="isSelf(openedChatUser?.id) ? '#10b981' : (openedChatUser?.avatar_color || '#6366f1')">
                <img *ngIf="openedChatUser?.avatar_url && !isSelf(openedChatUser?.id)" [src]="openedChatUser.avatar_url">
                <span *ngIf="!openedChatUser?.avatar_url && !isSelf(openedChatUser?.id)">{{ getInitials(openedChatUser?.name || '?') }}</span>
                <span *ngIf="isSelf(openedChatUser?.id)">ME</span>
            </div>
            <div style="display:flex; flex-direction:column; color:var(--header-text);">
                <span style="font-size:14px; font-weight:600;">{{ openedChatUser?.name }}</span>
                <span style="font-size:11px; color:var(--header-icon);">{{ isOnline(openedChatUser?.last_login) ? 'Online' : 'Away' }}</span>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:4px;">
            <span class="material-symbols-outlined icon-btn" [class.active]="selectionMode" (click)="toggleSelectionMode()" title="Select Messages">checklist</span>
            <span class="material-symbols-outlined icon-btn" (click)="closeChat()">remove</span>
            <span class="material-symbols-outlined icon-btn" (click)="close.emit()">close</span>
          </div>
        </div>
        <div class="selection-toolbar" *ngIf="selectionMode">
            <span style="font-weight:600; font-size:13px; color:var(--text-primary);">{{ selectedMessages.size }} selected</span>
            <div style="display:flex; gap:8px;">
                <button class="st-btn st-all" (click)="selectAllMessages()">All</button>
                <button class="st-btn st-cancel" (click)="toggleSelectionMode()">Cancel</button>
                <button class="st-btn st-delete" *ngIf="selectedMessages.size > 0" (click)="deleteSelectedMessages()">Delete</button>
                <button class="st-btn st-forward" *ngIf="selectedMessages.size > 0">Forward</button>
            </div>
        </div>
        <div class="chat-body-area" #scrollMe>
            <ng-container *ngFor="let msg of messages; let i = index">
                <div class="date-badge" *ngIf="isFirstOfDay(i, messages)" [id]="'date-badge-chat-' + i">
                    <span class="hover-date-badge" style="position:relative; cursor:pointer;" (click)="openCustomCalendar('chat')">
                        {{ getDateLabel(msg.created_at) }}
                    </span>
                </div>
                <div class="msg-row" [id]="'msg-' + msg.id" [class.mine]="msg.is_mine" [class.other]="!msg.is_mine">
                    <input type="checkbox" *ngIf="selectionMode && !msg.is_mine" [checked]="selectedMessages.has(msg.id)" (change)="toggleMessageSelection(msg)" style="margin-right:8px; margin-bottom:12px;">
                    <div style="display:flex; flex-direction:column; max-width:100%;">
                        <div *ngIf="!msg.is_file" class="msg-bubble" [ngClass]="{'emoji-only': isOnlyEmoji(msg.message)}">
                            <div *ngIf="getReplyMeta(msg.message) as rMsg" class="inline-reply">
                                <strong>{{ rMsg.sender_name || (rMsg.is_mine ? 'You' : (openedChatUser?.name || 'User')) }}</strong>
                                <span *ngIf="rMsg.is_file" style="display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:12px">description</span> {{ cleanMessage(rMsg.message) || 'Attachment' }}</span>
                                <span *ngIf="!rMsg.is_file" [innerHTML]="linkify(cleanMessage(rMsg.message))"></span>
                            </div>
                            <ng-container *ngIf="editingMessageId !== msg.id">
                                <ng-container *ngIf="msg.message.startsWith('[Sticker]') || msg.message.startsWith('[GIF]')">
                                    <div class="chat-sticker-bubble" style="background: transparent; padding: 0;">
                                        <img [src]="msg.message.startsWith('[Sticker]') ? msg.message.substring(9) : msg.message.substring(5)" class="chat-media-img" style="max-width: 150px; border-radius: 8px;" />
                                    </div>
                                </ng-container>
                                <ng-container *ngIf="!msg.message.startsWith('[Sticker]') && !msg.message.startsWith('[GIF]')">
                                    <span [innerHTML]="linkify(cleanMessage(msg.message))"></span>
                                </ng-container>
                            </ng-container>
                            <div *ngIf="editingMessageId === msg.id" style="margin-top: 4px; display:flex; gap:4px; flex-direction:column;">
                                <textarea [(ngModel)]="editingMessageText" class="edit-textarea" rows="2" style="width:100%; border:none; border-radius:4px; padding:4px; font-size:13px; color:#000"></textarea>
                                <div style="display:flex; justify-content:flex-end; gap:4px">
                                    <span style="font-size:12px; cursor:pointer; color:#ef4444" (click)="cancelEdit()">Cancel</span>
                                    <span style="font-size:12px; cursor:pointer; color:#22c55e" (click)="saveEditMessage(msg)">Save</span>
                                </div>
                            </div>
                            <div class="msg-hover-actions">
                                <ng-container *ngIf="deletingMessageId !== msg.id">
                                    <span class="material-symbols-outlined" title="React" (click)="reactToMessage(msg, $event)">add_reaction</span>
                                    <span class="material-symbols-outlined" title="Reply" (click)="replyToMessage(msg, $event)">reply</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Copy" (click)="copyMessage(msg, $event)">content_copy</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Edit" (click)="editMessage(msg, $event)">edit</span>
                                    <span class="material-symbols-outlined" title="Forward" (click)="forwardMessage(msg, $event)">forward</span>
                                    <span class="material-symbols-outlined" title="Delete" (click)="deleteMessage(msg, $event)">delete</span>
                                    <span class="material-symbols-outlined" title="Copy Link" (click)="copyLink(msg, $event)">link</span>
                                </ng-container>
                                <ng-container *ngIf="deletingMessageId === msg.id">
                                    <span style="font-size:12px; font-weight:600; color:var(--text-primary); margin-right:4px; white-space:nowrap; cursor:default;">Delete?</span>
                                    <div style="display:flex; gap:4px; margin-left:2px;">
                                        <button style="border:none; background:#ef4444; color:white; font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer; box-shadow:0 2px 4px rgba(239,68,68,0.3);" (click)="confirmDelete(msg, $event)">Yes</button>
                                        <button style="border:none; background:var(--border-color); color:var(--text-primary); font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer;" (click)="cancelDelete($event)">No</button>
                                    </div>
                                </ng-container>
                            </div>
                            <div class="msg-reaction-popover" *ngIf="reactingMessageId === msg.id">
                                <span *ngFor="let emoji of ['👍','❤️','😂','😲','😢','😡']" (click)="addReaction(msg, emoji); $event.stopPropagation()">{{ emoji }}</span>
                                <span class="material-symbols-outlined" style="font-size:14px; cursor:pointer;" (click)="closeReactionPopover($event)">close</span>
                            </div>
                            <div class="msg-reactions-display" *ngIf="getReactionsList(msg).length > 0">
                                <span *ngFor="let r of getReactionsList(msg)" (click)="addReaction(msg, r.split(' ')[0])">{{ r }}</span>
                            </div>
                        </div>
                        <div *ngIf="msg.is_file" (click)="openFile(msg.file_path, msg.message, msg, false)" [ngClass]="isImage(msg.message || msg.file_path) ? 'msg-image-wrapper' : 'msg-file-card'">
                            <img *ngIf="isImage(msg.message || msg.file_path)" [src]="getFileUrl(msg.file_path)" class="chat-image-preview" (error)="$event.stopPropagation()">
                            <ng-container *ngIf="!isImage(msg.message || msg.file_path)">
                                <div class="ac-icon">
                                    <span *ngIf="!msg.isUploading" class="material-symbols-outlined">description</span>
                                    <div *ngIf="msg.isUploading" class="upload-spinner"></div>
                                </div>
                                <span class="ac-name">{{ cleanMessage(msg.message) || 'Attachment' }}</span>
                            </ng-container>
                            <div class="msg-hover-actions">
                                <ng-container *ngIf="deletingMessageId !== msg.id">
                                    <span class="material-symbols-outlined" title="React" (click)="reactToMessage(msg, $event)">add_reaction</span>
                                    <span class="material-symbols-outlined" title="Reply" (click)="replyToMessage(msg, $event)">reply</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Copy" (click)="copyMessage(msg, $event)">content_copy</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Edit" (click)="editMessage(msg, $event)">edit</span>
                                    <span class="material-symbols-outlined" title="Forward" (click)="forwardMessage(msg, $event)">forward</span>
                                    <span class="material-symbols-outlined" title="Delete" (click)="deleteMessage(msg, $event)">delete</span>
                                    <span class="material-symbols-outlined" title="Download" (click)="downloadMessage(msg, $event)">download</span>
                                </ng-container>
                                <ng-container *ngIf="deletingMessageId === msg.id">
                                    <span style="font-size:12px; font-weight:600; color:var(--text-primary); margin-right:4px; white-space:nowrap; cursor:default;">Delete?</span>
                                    <div style="display:flex; gap:4px; margin-left:2px;">
                                        <button style="border:none; background:#ef4444; color:white; font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer; box-shadow:0 2px 4px rgba(239,68,68,0.3);" (click)="confirmDelete(msg, $event)">Yes</button>
                                        <button style="border:none; background:var(--border-color); color:var(--text-primary); font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer;" (click)="cancelDelete($event)">No</button>
                                    </div>
                                </ng-container>
                            </div>
                            <div class="msg-reaction-popover" *ngIf="reactingMessageId === msg.id">
                                <span *ngFor="let emoji of ['👍','❤️','😂','😲','😢','😡']" (click)="addReaction(msg, emoji); $event.stopPropagation()">{{ emoji }}</span>
                                <span class="material-symbols-outlined" style="font-size:14px; cursor:pointer;" (click)="closeReactionPopover($event)">close</span>
                            </div>
                            <div class="msg-reactions-display" *ngIf="getReactionsList(msg).length > 0">
                                <span *ngFor="let r of getReactionsList(msg)" (click)="addReaction(msg, r.split(' ')[0])">{{ r }}</span>
                            </div>
                        </div>
                        <div class="msg-meta">
                            <span *ngIf="msg.is_edited && editingMessageId !== msg.id" style="font-size:10px; font-style:italic; margin-right:4px;">(edited)</span>
                            {{ formatTime(msg.created_at) }}
                            <span *ngIf="msg.is_mine" class="material-symbols-outlined" style="font-size:14px; margin-left:4px;" [style.color]="(msg.is_read || isSelf(openedChatUser?.id)) ? '#3b82f6' : 'inherit'">{{ (msg.is_read || isSelf(openedChatUser?.id)) ? 'done_all' : 'check' }}</span>
                        </div>
                    </div>
                    <input type="checkbox" *ngIf="selectionMode && msg.is_mine" [checked]="selectedMessages.has(msg.id)" (change)="toggleMessageSelection(msg)" style="margin-left:8px; margin-bottom:12px;">
                </div>
            </ng-container>
            <div style="height: 24px; flex-shrink: 0; width: 100%;"></div>
        </div>
        <div class="chat-input-container">
           <div *ngIf="stagedFile" class="staged-file">
               <span class="material-symbols-outlined" style="font-size:16px;">attach_file</span>
               <span class="staged-name">{{ stagedFile.name }}</span>
               <span class="material-symbols-outlined close-btn" (click)="removeStagedFile()">close</span>
           </div>
           <div *ngIf="replyingTo" class="reply-banner">
               <span class="material-symbols-outlined" style="margin-right:8px; font-size:20px; color:#3b82f6">reply</span>
               <div class="reply-content">
                   <strong>{{ replyingTo.sender_name || (replyingTo.is_mine ? 'You' : (openedChatUser?.name || 'User')) }}</strong>
                   <span *ngIf="replyingTo.is_file" style="display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:12px">description</span> {{ cleanMessage(replyingTo.message) || 'Attachment' }}</span>
                   <span *ngIf="!replyingTo.is_file" [innerHTML]="linkify(cleanMessage(replyingTo.message))"></span>
               </div>
               <span class="material-symbols-outlined" (click)="cancelReply()">close</span>
           </div>
           <div class="chat-input-area">
               <input type="file" #fileInputChat style="display:none" (change)="onFileSelected($event)">
               <span class="material-symbols-outlined icon-btn" (click)="fileInputChat.click()">attach_file</span>
               <span class="material-symbols-outlined icon-btn emoji-toggle-btn" [class.active]="showEmojiPicker" (click)="$event.stopPropagation(); showEmojiPicker = !showEmojiPicker">sentiment_satisfied</span>
               <textarea class="chat-input-textarea" placeholder="Message {{ openedChatUser?.name?.split(' ')[0] }}..." [(ngModel)]="newMessage" (keydown)="onInputKeydown($event)" (input)="autoResizeInput($event)" (paste)="onPaste($event)" rows="1"></textarea>
               <button class="send-btn" (click)="sendMessage()">
                   <span class="material-symbols-outlined" style="font-size:18px; margin-left:2px;">send</span>
               </button>
           </div>
           <app-media-picker *ngIf="showEmojiPicker" class="media-picker-inline" (click)="$event.stopPropagation()" (emojiSelect)="addEmoji($event)" (mediaSelect)="sendMedia($event)"></app-media-picker>
        </div>
      </ng-container>

      <!-- CHANNELS VIEW -->
      <ng-container *ngIf="activeWidget === 'channels' && !openedChannelId">
        <div class="wp-header-white" style="background:var(--bg-color); color:var(--text-primary); border-bottom:1px solid var(--border-color);">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined" style="color:var(--ac-icon-color);">groups</span>
            <span style="font-size:15px; font-weight:600;">Groups</span>
          </div>
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined icon-btn" (click)="showCreateGroup = !showCreateGroup">group_add</span>
            <span class="material-symbols-outlined icon-btn" (click)="close.emit()">close</span>
          </div>
        </div>
        
        <div *ngIf="showCreateGroup" class="create-group-form" style="padding: 12px; background: var(--bg-color); border-bottom: 1px solid var(--border-color);">
            <input type="text" placeholder="Group name..." [(ngModel)]="newGroupName" style="width: 100%; padding: 8px 12px; border: 1px solid var(--border-color); border-radius: 6px; margin-bottom: 8px; background: var(--bg-color); color: var(--text-primary); font-family: inherit; font-size: 13px; box-sizing: border-box;">
            <input type="text" placeholder="Description (optional)" [(ngModel)]="newGroupDescription" style="width: 100%; padding: 8px 12px; border: 1px solid var(--border-color); border-radius: 6px; margin-bottom: 8px; background: var(--bg-color); color: var(--text-primary); font-family: inherit; font-size: 13px; box-sizing: border-box;">
            <div style="margin-bottom: 12px; display: flex; align-items: center;">
                <button (click)="fileInputGroupAvatar.click()" style="padding: 6px 12px; border: 1px solid var(--border-color); border-radius: 6px; background: #f3f4f6; color: #374151; cursor: pointer; font-size: 13px;">Choose Avatar...</button>
                <input type="file" #fileInputGroupAvatar style="display:none" (change)="onGroupAvatarSelected($event)" accept="image/*">
                <span *ngIf="newGroupAvatar" style="margin-left: 8px; font-size: 12px; color: var(--text-secondary); max-width: 120px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{{ newGroupAvatar.name }}</span>
            </div>
            <div style="display: flex; justify-content: flex-end; gap: 8px;">
                <button (click)="showCreateGroup = false; resetCreateGroup()" style="padding: 6px 16px; border: 1px solid var(--border-color); border-radius: 6px; background: white; color: #374151; cursor: pointer; font-size: 13px; font-weight: 500;">Cancel</button>
                <button (click)="createGroup()" [disabled]="!newGroupName.trim()" style="padding: 6px 16px; border: none; border-radius: 6px; background: #60a5fa; color: white; cursor: pointer; font-size: 13px; font-weight: 500;" [style.opacity]="!newGroupName.trim() ? '0.5' : '1'">Create</button>
            </div>
        </div>
        <div class="wp-search">
          <span class="material-symbols-outlined">search</span>
          <input type="text" placeholder="Search groups..." [(ngModel)]="searchQuery" (ngModelChange)="onGroupSearchChange($event)">
        </div>
        <div class="wp-section-title" *ngIf="filteredChannels().length > 0 || !searchQuery">MY GROUPS</div>
        <div class="wp-body-list" *ngIf="filteredChannels().length > 0 || !searchQuery">
          <div class="chat-list-item" *ngFor="let c of filteredChannels()" (click)="openChannel(c.id, c.name, c.avatar_url)">
            <div class="cli-av" style="background:var(--bg-color); color:var(--text-primary); border:1px solid var(--border-color); overflow:hidden; padding:0;">
              <img *ngIf="c.avatar_url" [src]="c.avatar_url" style="width:100%; height:100%; object-fit:cover;">
              <span *ngIf="!c.avatar_url" style="font-size:16px; font-weight:bold;">#</span>
            </div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ c.name }}</span>
                <span class="cli-time">{{ formatTime(c.last_time) }}</span>
              </div>
              <div class="cli-preview" [innerHTML]="linkify(cleanMessage(c.last_message)) || 'No messages'"></div>
            </div>
          </div>
          <div *ngIf="channels.length === 0 && !searchQuery" style="padding: 20px; text-align:center; color:var(--text-secondary); font-size:13px;">
            No channels
          </div>
        </div>
        
        <div class="wp-section-title" *ngIf="searchQuery && filteredServerGroups().length > 0" style="margin-top:8px;">ALL GROUPS</div>
        <div class="wp-body-list" *ngIf="searchQuery && filteredServerGroups().length > 0">
          <div class="chat-list-item" *ngFor="let c of filteredServerGroups()" (click)="openChannel(c.id, c.name, c.avatar_url)">
            <div class="cli-av" style="background:var(--bg-color); color:var(--text-primary); border:1px solid var(--border-color); overflow:hidden; padding:0;">
              <img *ngIf="c.avatar_url" [src]="c.avatar_url" style="width:100%; height:100%; object-fit:cover;">
              <span *ngIf="!c.avatar_url" style="font-size:16px; font-weight:bold;">#</span>
            </div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ c.name }}</span>
              </div>
              <div class="cli-preview" style="color:var(--text-secondary)">Channel</div>
            </div>
          </div>
        </div>
      </ng-container>

      <!-- OPENED CHANNEL VIEW -->
      <ng-container *ngIf="activeWidget === 'channels' && openedChannelId">
        <div class="wp-header-chat">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined icon-btn" (click)="closeChannel()">arrow_back</span>
            <div class="cli-av" style="width:32px; height:32px; font-size:16px; margin-right:4px; background:var(--bg-color); color:var(--text-primary); border:1px solid var(--border-color); overflow:hidden; padding:0;">
              <img *ngIf="openedChannelAvatarUrl" [src]="openedChannelAvatarUrl" style="width:100%; height:100%; object-fit:cover;">
              <span *ngIf="!openedChannelAvatarUrl" style="font-size:16px; font-weight:bold;">#</span>
            </div>
            <div style="display:flex; flex-direction:column; color:var(--header-text); cursor:pointer;" (click)="openGroupInfo()">
                <span style="font-size:14px; font-weight:600;">{{ openedChannelName }}</span>
                <span style="font-size:11px; color:var(--header-icon);">{{ channelMembers.length }} members</span>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:4px;">
            <span class="material-symbols-outlined icon-btn" [class.active]="selectionMode" (click)="toggleSelectionMode()" title="Select Messages">checklist</span>
            <span class="material-symbols-outlined icon-btn" *ngIf="isAdmin" (click)="openAddMember()">person_add</span>
            <span class="material-symbols-outlined icon-btn" style="color:#ef4444;" (click)="leaveGroup()">logout</span>
            <span class="material-symbols-outlined icon-btn" (click)="closeChannel()">remove</span>
            <span class="material-symbols-outlined icon-btn" (click)="close.emit()">close</span>
          </div>
        </div>
        <div class="selection-toolbar" *ngIf="selectionMode">
            <span style="font-weight:600; font-size:13px; color:var(--text-primary);">{{ selectedMessages.size }} selected</span>
            <div style="display:flex; gap:8px;">
                <button class="st-btn st-all" (click)="selectAllMessages()">All</button>
                <button class="st-btn st-cancel" (click)="toggleSelectionMode()">Cancel</button>
                <button class="st-btn st-delete" *ngIf="selectedMessages.size > 0" (click)="deleteSelectedMessages()">Delete</button>
                <button class="st-btn st-forward" *ngIf="selectedMessages.size > 0">Forward</button>
            </div>
        </div>
        <div class="chat-body-area" #scrollMe>
            <ng-container *ngFor="let msg of channelMessages; let i = index">
                <div class="date-badge" *ngIf="isFirstOfDay(i, channelMessages)" [id]="'date-badge-channel-' + i">
                    <span class="hover-date-badge" style="position:relative; cursor:pointer;" (click)="openCustomCalendar('channel')">
                        {{ getDateLabel(msg.created_at) }}
                    </span>
                </div>
                <div class="msg-row" [id]="'msg-' + msg.id" [class.mine]="msg.is_mine" [class.other]="!msg.is_mine">
                    <input type="checkbox" *ngIf="selectionMode && !msg.is_mine" [checked]="selectedMessages.has(msg.id)" (change)="toggleMessageSelection(msg)" style="margin-right:8px; margin-bottom:12px;">
                    <div style="display:flex; flex-direction:column; max-width:100%;">
                        <span *ngIf="!msg.is_mine" style="font-size:11px; color:var(--text-secondary); margin-bottom:2px; margin-left:4px;">{{ msg.sender_name }}</span>
                        <div *ngIf="!msg.is_file" class="msg-bubble" [ngClass]="{'emoji-only': isOnlyEmoji(msg.message)}">
                            <div *ngIf="getReplyMeta(msg.message) as rMsg" class="inline-reply">
                                <strong>{{ rMsg.sender_name || (rMsg.is_mine ? 'You' : 'User') }}</strong>
                                <span *ngIf="rMsg.is_file" style="display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:12px">description</span> {{ cleanMessage(rMsg.message) || 'Attachment' }}</span>
                                <span *ngIf="!rMsg.is_file" [innerHTML]="linkify(cleanMessage(rMsg.message))"></span>
                            </div>
                            <ng-container *ngIf="editingMessageId !== msg.id">
                                <ng-container *ngIf="msg.message.startsWith('[Sticker]') || msg.message.startsWith('[GIF]')">
                                    <div class="chat-sticker-bubble" style="background: transparent; padding: 0;">
                                        <img [src]="msg.message.startsWith('[Sticker]') ? msg.message.substring(9) : msg.message.substring(5)" class="chat-media-img" style="max-width: 150px; border-radius: 8px;" />
                                    </div>
                                </ng-container>
                                <ng-container *ngIf="!msg.message.startsWith('[Sticker]') && !msg.message.startsWith('[GIF]')">
                                    <span [innerHTML]="linkify(cleanMessage(msg.message))"></span>
                                </ng-container>
                            </ng-container>
                            <div *ngIf="editingMessageId === msg.id" style="margin-top: 4px; display:flex; gap:4px; flex-direction:column;">
                                <textarea [(ngModel)]="editingMessageText" class="edit-textarea" rows="2" style="width:100%; border:none; border-radius:4px; padding:4px; font-size:13px; color:#000"></textarea>
                                <div style="display:flex; justify-content:flex-end; gap:4px">
                                    <span style="font-size:12px; cursor:pointer; color:#ef4444" (click)="cancelEdit()">Cancel</span>
                                    <span style="font-size:12px; cursor:pointer; color:#22c55e" (click)="saveEditMessage(msg)">Save</span>
                                </div>
                            </div>
                            <div class="msg-hover-actions">
                                <ng-container *ngIf="deletingMessageId !== msg.id">
                                    <span class="material-symbols-outlined" title="React" (click)="reactToMessage(msg, $event)">add_reaction</span>
                                    <span class="material-symbols-outlined" title="Reply" (click)="replyToMessage(msg, $event)">reply</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Copy" (click)="copyMessage(msg, $event)">content_copy</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Edit" (click)="editMessage(msg, $event)">edit</span>
                                    <span class="material-symbols-outlined" title="Forward" (click)="forwardMessage(msg, $event)">forward</span>
                                    <span class="material-symbols-outlined" title="Delete" (click)="deleteMessage(msg, $event)">delete</span>
                                    <span class="material-symbols-outlined" title="Copy Link" (click)="copyLink(msg, $event)">link</span>
                                </ng-container>
                                <ng-container *ngIf="deletingMessageId === msg.id">
                                    <span style="font-size:12px; font-weight:600; color:var(--text-primary); margin-right:4px; white-space:nowrap; cursor:default;">Delete?</span>
                                    <div style="display:flex; gap:4px; margin-left:2px;">
                                        <button style="border:none; background:#ef4444; color:white; font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer; box-shadow:0 2px 4px rgba(239,68,68,0.3);" (click)="confirmDelete(msg, $event)">Yes</button>
                                        <button style="border:none; background:var(--border-color); color:var(--text-primary); font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer;" (click)="cancelDelete($event)">No</button>
                                    </div>
                                </ng-container>
                            </div>
                            <div class="msg-reaction-popover" *ngIf="reactingMessageId === msg.id">
                                <span *ngFor="let emoji of ['👍','❤️','😂','😲','😢','😡']" (click)="addReaction(msg, emoji); $event.stopPropagation()">{{ emoji }}</span>
                                <span class="material-symbols-outlined" style="font-size:14px; cursor:pointer;" (click)="closeReactionPopover($event)">close</span>
                            </div>
                            <div class="msg-reactions-display" *ngIf="getReactionsList(msg).length > 0">
                                <span *ngFor="let r of getReactionsList(msg)" (click)="addReaction(msg, r.split(' ')[0])">{{ r }}</span>
                            </div>
                        </div>
                        <div *ngIf="msg.is_file" (click)="openFile(msg.file_path, msg.message, msg, true)" [ngClass]="isImage(msg.message || msg.file_path) ? 'msg-image-wrapper' : 'msg-file-card'">
                            <img *ngIf="isImage(msg.message || msg.file_path)" [src]="getFileUrl(msg.file_path)" class="chat-image-preview" (error)="$event.stopPropagation()">
                            <ng-container *ngIf="!isImage(msg.message || msg.file_path)">
                                <div class="ac-icon">
                                    <span *ngIf="!msg.isUploading" class="material-symbols-outlined">description</span>
                                    <div *ngIf="msg.isUploading" class="upload-spinner"></div>
                                </div>
                                <span class="ac-name">{{ cleanMessage(msg.message) || 'Attachment' }}</span>
                            </ng-container>
                            <div class="msg-hover-actions">
                                <ng-container *ngIf="deletingMessageId !== msg.id">
                                    <span class="material-symbols-outlined" title="React" (click)="reactToMessage(msg, $event)">add_reaction</span>
                                    <span class="material-symbols-outlined" title="Reply" (click)="replyToMessage(msg, $event)">reply</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Copy" (click)="copyMessage(msg, $event)">content_copy</span>
                                    <span *ngIf="!msg.message?.startsWith('[Sticker]') && !msg.message?.startsWith('[GIF]')" class="material-symbols-outlined" title="Edit" (click)="editMessage(msg, $event)">edit</span>
                                    <span class="material-symbols-outlined" title="Forward" (click)="forwardMessage(msg, $event)">forward</span>
                                    <span class="material-symbols-outlined" title="Delete" (click)="deleteMessage(msg, $event)">delete</span>
                                    <span class="material-symbols-outlined" title="Download" (click)="downloadMessage(msg, $event)">download</span>
                                </ng-container>
                                <ng-container *ngIf="deletingMessageId === msg.id">
                                    <span style="font-size:12px; font-weight:600; color:var(--text-primary); margin-right:4px; white-space:nowrap; cursor:default;">Delete?</span>
                                    <div style="display:flex; gap:4px; margin-left:2px;">
                                        <button style="border:none; background:#ef4444; color:white; font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer; box-shadow:0 2px 4px rgba(239,68,68,0.3);" (click)="confirmDelete(msg, $event)">Yes</button>
                                        <button style="border:none; background:var(--border-color); color:var(--text-primary); font-size:11px; font-weight:600; padding:2px 10px; border-radius:12px; cursor:pointer;" (click)="cancelDelete($event)">No</button>
                                    </div>
                                </ng-container>
                            </div>
                            <div class="msg-reaction-popover" *ngIf="reactingMessageId === msg.id">
                                <span *ngFor="let emoji of ['👍','❤️','😂','😲','😢','😡']" (click)="addReaction(msg, emoji); $event.stopPropagation()">{{ emoji }}</span>
                                <span class="material-symbols-outlined" style="font-size:14px; cursor:pointer;" (click)="closeReactionPopover($event)">close</span>
                            </div>
                            <div class="msg-reactions-display" *ngIf="getReactionsList(msg).length > 0">
                                <span *ngFor="let r of getReactionsList(msg)" (click)="addReaction(msg, r.split(' ')[0])">{{ r }}</span>
                            </div>
                        </div>
                        <div class="msg-meta">
                            <span *ngIf="msg.is_edited && editingMessageId !== msg.id" style="font-size:10px; font-style:italic; margin-right:4px;">(edited)</span>
                            {{ formatTime(msg.created_at) }}
                            <span *ngIf="msg.is_mine" class="material-symbols-outlined" style="font-size:14px; margin-left:4px;" [style.color]="(msg.is_read || isSelf(openedChatUser?.id)) ? '#3b82f6' : 'inherit'">{{ (msg.is_read || isSelf(openedChatUser?.id)) ? 'done_all' : 'check' }}</span>
                        </div>
                    </div>
                    <input type="checkbox" *ngIf="selectionMode && msg.is_mine" [checked]="selectedMessages.has(msg.id)" (change)="toggleMessageSelection(msg)" style="margin-left:8px; margin-bottom:12px;">
                </div>
            </ng-container>
            <div style="height: 24px; flex-shrink: 0; width: 100%;"></div>
        </div>
        <div class="chat-input-container">
           <div *ngIf="stagedFile" class="staged-file">
               <span class="material-symbols-outlined" style="font-size:16px;">attach_file</span>
               <span class="staged-name">{{ stagedFile.name }}</span>
               <span class="material-symbols-outlined close-btn" (click)="removeStagedFile()">close</span>
           </div>
           <div *ngIf="replyingTo" class="reply-banner">
               <span class="material-symbols-outlined" style="margin-right:8px; font-size:20px; color:#3b82f6">reply</span>
               <div class="reply-content">
                   <strong>{{ replyingTo.sender_name || (replyingTo.is_mine ? 'You' : 'User') }}</strong>
                   <span *ngIf="replyingTo.is_file" style="display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:12px">description</span> {{ cleanMessage(replyingTo.message) || 'Attachment' }}</span>
                   <span *ngIf="!replyingTo.is_file" [innerHTML]="linkify(cleanMessage(replyingTo.message))"></span>
               </div>
               <span class="material-symbols-outlined" (click)="cancelReply()">close</span>
           </div>
           <div class="chat-input-area">
               <input type="file" #fileInputChannel style="display:none" (change)="onFileSelected($event)">
               <span class="material-symbols-outlined icon-btn" (click)="fileInputChannel.click()">attach_file</span>
               <span class="material-symbols-outlined icon-btn emoji-toggle-btn" [class.active]="showEmojiPicker" (click)="$event.stopPropagation(); showEmojiPicker = !showEmojiPicker">sentiment_satisfied</span>
               <textarea class="chat-input-textarea" placeholder="Message {{ openedChannelName }}..." [(ngModel)]="newMessage" (keydown)="onInputKeydown($event)" (input)="autoResizeInput($event)" (paste)="onPaste($event)" rows="1"></textarea>
               <button class="send-btn" (click)="sendMessage()">
                   <span class="material-symbols-outlined" style="font-size:18px; margin-left:2px;">send</span>
               </button>
           </div>
           <app-media-picker *ngIf="showEmojiPicker" class="media-picker-inline" (click)="$event.stopPropagation()" (emojiSelect)="addEmoji($event)" (mediaSelect)="sendMedia($event)"></app-media-picker>
        </div>
      </ng-container>

      <!-- Forward Modal -->
      <div class="forward-modal-overlay" *ngIf="showForwardModal" (click)="closeForwardModal()">
          <div class="forward-modal-content" (click)="$event.stopPropagation()">
              <div class="fm-header">
                  <span class="material-symbols-outlined" style="color:#3b82f6; margin-right:8px;">forward</span>
                  <h3>Forward messages to</h3>
                  <div style="flex:1"></div>
                  <span class="material-symbols-outlined" style="cursor:pointer;" (click)="closeForwardModal()">close</span>
              </div>
              <div class="fm-search">
                  <span class="material-symbols-outlined">search</span>
                  <input type="text" placeholder="Search recipients" [(ngModel)]="forwardSearch">
              </div>
              <div class="fm-list">
                  <div class="fm-item" *ngFor="let u of filteredForwardUsers()" (click)="toggleForwardSelection('user_' + u.id)">
                      <div class="fm-av" [style.background]="u.avatar_color">{{ (u.name || '').charAt(0).toUpperCase() }}</div>
                      <div class="fm-info">
                          <strong>{{ u.name }} <span *ngIf="isSelf(u.id)">(You)</span></strong>
                          <span>{{ u.email }}</span>
                      </div>
                      <input type="checkbox" [checked]="forwardSelection.has('user_' + u.id)" (click)="$event.stopPropagation(); toggleForwardSelection('user_' + u.id)">
                  </div>
                  <div class="fm-item" *ngFor="let c of filteredForwardChannels()" (click)="toggleForwardSelection('channel_' + c.id)">
                      <div class="fm-av" style="background:#4f46e5;">#</div>
                      <div class="fm-info">
                          <strong>{{ c.name }}</strong>
                          <span>Channel</span>
                      </div>
                      <input type="checkbox" [checked]="forwardSelection.has('channel_' + c.id)" (click)="$event.stopPropagation(); toggleForwardSelection('channel_' + c.id)">
                  </div>
              </div>
              <div class="fm-footer">
                  <span>{{ forwardSelection.size }} recipients selected</span>
                  <button class="send-btn" style="width:auto; padding: 0 16px; border-radius: 4px;" (click)="sendForward()" [disabled]="forwardSelection.size === 0">Send</button>
              </div>
          </div>
      </div>
      <!-- Toast Overlay -->
      <div class="widget-toast" *ngIf="toastMessage">
          {{ toastMessage }}
      </div>
      <!-- Image Preview Modal -->
      <div class="image-preview-overlay" *ngIf="previewImageUrl" (click)="closePreview()">
          <div class="image-preview-header" (click)="$event.stopPropagation()">
              <div class="preview-actions">
                  <span class="material-symbols-outlined action-icon" (click)="zoomPreview(-0.2)">zoom_out</span>
                  <span class="material-symbols-outlined action-icon" (click)="zoomPreview(0.2)">zoom_in</span>
                  <div class="action-btn download-btn" (click)="downloadPreviewImage()">
                      <span class="material-symbols-outlined" style="color:#3b82f6; font-size: 20px;">download</span>
                  </div>
                  <div class="action-btn close-btn" (click)="closePreview()">
                      <span class="material-symbols-outlined" style="color:#ef4444; font-size: 20px;">close</span>
                  </div>
              </div>
          </div>
          
          <div class="preview-nav-left" *ngIf="hasPrevImage()" (click)="prevImage(); $event.stopPropagation()">
              <span class="material-symbols-outlined">navigate_before</span>
          </div>
          
          <div class="image-preview-content" (click)="$event.stopPropagation()" (wheel)="onPreviewWheel($event)">
              <img [src]="previewImageUrl" class="preview-full-image" [style.transform]="'scale(' + previewZoomLevel + ')'">
          </div>
          
          <div class="preview-nav-right" *ngIf="hasNextImage()" (click)="nextImage(); $event.stopPropagation()">
              <span class="material-symbols-outlined">navigate_next</span>
          </div>
      </div>
        <!-- CUSTOM CALENDAR MODAL -->
        <div class="calendar-modal-overlay" *ngIf="showCalendarModal" (click)="showCalendarModal = false">
            <div class="calendar-modal" (click)="$event.stopPropagation()">
                <div class="cal-header">
                    <button class="cal-nav-btn" (click)="changeMonth(-1)"><span class="material-symbols-outlined" style="font-size:20px;">chevron_left</span></button>
                    <span class="cal-month-title">{{ calendarCurrentMonth | date:'MMMM yyyy' }}</span>
                    <button class="cal-nav-btn" (click)="changeMonth(1)"><span class="material-symbols-outlined" style="font-size:20px;">chevron_right</span></button>
                </div>
                <div class="cal-weekdays">
                    <span>Su</span><span>Mo</span><span>Tu</span><span>We</span><span>Th</span><span>Fr</span><span>Sa</span>
                </div>
                <div class="cal-grid">
                    <div class="cal-cell" *ngFor="let d of calendarDays" 
                         [class.empty]="!d" 
                         [class.has-msg]="d && hasMessagesOnDate(d)"
                         (click)="d && selectCustomDate(d)">
                        <span *ngIf="d">{{ d.getDate() }}</span>
                        <div class="cal-dot" *ngIf="d && hasMessagesOnDate(d)"></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Group Info Modal -->
        <div class="group-modal-overlay" *ngIf="showGroupInfoModal" (click)="showGroupInfoModal = false">
            <div class="group-modal" (click)="$event.stopPropagation()">
                <div class="gm-header">
                    <span class="material-symbols-outlined" style="color:#3b82f6; font-size: 20px;">info</span>
                    <h3 style="margin: 0; font-size: 15px; font-weight: 600; color: var(--text-primary);">Group Info</h3>
                    <div style="flex:1"></div>
                    <span class="material-symbols-outlined icon-btn" style="cursor:pointer;" (click)="showGroupInfoModal = false">close</span>
                </div>
                <div class="gm-body">
                    <div *ngIf="isAdmin" style="margin-bottom: 16px;">
                        <label style="font-size: 11px; font-weight: 600; color: var(--text-secondary); margin-bottom: 4px; display: block;">Group Name</label>
                        <input type="text" [(ngModel)]="editGroupName" style="width: 100%; padding: 8px 12px; border: 1px solid var(--input-border); border-radius: 6px; margin-bottom: 12px; background: var(--input-bg); color: var(--text-primary); font-family: inherit; font-size: 13px; box-sizing: border-box; outline: none; transition: border-color 0.2s;">
                        
                        <label style="font-size: 11px; font-weight: 600; color: var(--text-secondary); margin-bottom: 4px; display: block;">Description</label>
                        <input type="text" [(ngModel)]="editGroupDesc" style="width: 100%; padding: 8px 12px; border: 1px solid var(--input-border); border-radius: 6px; margin-bottom: 16px; background: var(--input-bg); color: var(--text-primary); font-family: inherit; font-size: 13px; box-sizing: border-box; outline: none; transition: border-color 0.2s;">
                        
                        <button class="gm-save-btn" (click)="saveGroupInfo()">Save Changes</button>
                    </div>
                    <div *ngIf="!isAdmin" style="margin-bottom: 16px;">
                        <div style="font-weight:700; font-size:15px; color: var(--text-primary); margin-bottom:4px;">#{{ openedChannelName }}</div>
                        <div style="font-size:13px; color: var(--text-secondary);">{{ editGroupDesc || 'No description' }}</div>
                    </div>

                    <hr style="border:none; border-top:1px solid var(--border-color); margin:16px 0;" />
                    
                    <div style="margin-bottom: 12px; font-size: 13px; font-weight: 700; color: var(--text-primary);">
                        Members ({{ channelMembers.length }})
                    </div>
                    <div class="gm-members-list" style="max-height: 200px; overflow-y: auto; display:flex; flex-direction:column; gap:12px;">
                        <div class="gm-member-item" *ngFor="let m of channelMembers" style="display:flex; align-items:center; gap:12px; font-size:13px;">
                            <div class="gm-member-av" [style.background]="m.avatar_color || '#6366f1'" style="width:28px; height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-size:11px; font-weight:600; overflow:hidden; flex-shrink:0;">
                                <img *ngIf="m.avatar_url" [src]="m.avatar_url" style="width:100%;height:100%;object-fit:cover;">
                                <span *ngIf="!m.avatar_url">{{ getInitials(m.name) }}</span>
                            </div>
                            <div class="gm-member-info" style="display:flex; align-items:center; flex:1; min-width:0;">
                                <span class="gm-member-name" style="color: var(--text-primary); font-weight:500; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{{ m.name }}</span>
                                <span *ngIf="m.role === 'admin' || m.is_admin" class="gm-role-badge" style="margin-left:6px; font-size:10px; background:rgba(99, 102, 241, 0.1); color:#6366f1; border: 1px solid rgba(99, 102, 241, 0.2); padding:2px 6px; border-radius:4px; font-weight:600;">Admin</span>
                            </div>
                            <div *ngIf="isAdmin && m.id !== currentUser?.id" style="display:flex; gap:8px; align-items:center;">
                                <button *ngIf="m.role !== 'admin' && !m.is_admin" (click)="setMemberAdmin(m.id, true)" style="background:none; border:none; color:#818cf8; cursor:pointer; font-size:12px; font-weight:500;">Make Admin</button>
                                <button *ngIf="m.role === 'admin' || m.is_admin" (click)="setMemberAdmin(m.id, false)" style="background:none; border:none; color:var(--text-secondary); cursor:pointer; font-size:12px; font-weight:500;">Remove Admin</button>
                                <span class="material-symbols-outlined rm-member-btn" (click)="removeMember(m.id)" title="Remove" style="font-size:18px; color:#ef4444; cursor:pointer;">person_remove</span>
                            </div>
                        </div>
                    </div>
                    
                    <button class="gm-delete-btn" (click)="deleteGroup()" *ngIf="isAdmin" style="width:100%; background:rgba(239,68,68,0.1); color:#ef4444; border:1px solid rgba(239,68,68,0.3); padding:10px; border-radius:8px; font-weight:600; cursor:pointer; margin-top:16px; transition:0.2s;">Delete Group For Everyone</button>
                </div>
            </div>
        </div>

        <!-- Add Member Modal -->
        <div class="group-modal-overlay" *ngIf="showAddMemberModal" (click)="showAddMemberModal = false">
            <div class="group-modal" (click)="$event.stopPropagation()" style="max-height: 80vh; display:flex; flex-direction: column;">
                <div class="gm-header">
                    <span class="material-symbols-outlined" style="color:#3b82f6; font-size: 20px;">person_add</span>
                    <h3 style="margin: 0; font-size: 15px; font-weight: 600; color: var(--text-primary);">Add to #{{ openedChannelName }}</h3>
                    <div style="flex:1"></div>
                    <span class="material-symbols-outlined icon-btn" style="cursor:pointer;" (click)="showAddMemberModal = false">close</span>
                </div>
                <div class="gm-search">
                    <span class="material-symbols-outlined">search</span>
                    <input type="text" placeholder="Search contacts..." [(ngModel)]="addMemberSearchQuery" (ngModelChange)="onAddMemberSearchChange()">
                </div>
                <div class="gm-add-list" style="overflow-y: auto; flex: 1; padding: 12px 16px; display:flex; flex-direction:column; gap:8px;">
                    <div class="gm-add-item" *ngFor="let c of filteredAddMemberContacts()" style="display:flex; align-items:center; gap:12px;">
                        <div class="gm-member-av" [style.background]="c.avatar_color || '#6366f1'" style="width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-size:12px; font-weight:600; overflow:hidden;">
                            <img *ngIf="c.avatar_url" [src]="c.avatar_url" style="width:100%;height:100%;object-fit:cover;">
                            <span *ngIf="!c.avatar_url">{{ getInitials(c.name) }}</span>
                        </div>
                        <div class="gm-add-info" style="flex:1; min-width:0;">
                            <span class="gm-add-name" style="color: var(--text-primary); font-size:13px; font-weight:500; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; display:block;">{{ c.name }}</span>
                        </div>
                        <button class="gm-add-btn" (click)="addMember(c.id)" style="padding:4px 10px; font-size:12px; background:rgba(59, 130, 246, 0.1); color:#3b82f6; border:1px solid rgba(59, 130, 246, 0.2); border-radius:6px; cursor:pointer; font-weight:600;">Add</button>
                    </div>
                    <div *ngIf="filteredAddMemberContacts().length === 0" style="text-align:center; padding: 20px; color: var(--text-secondary); font-size: 13px;">
                        No contacts found or all are already members.
                    </div>
                </div>
            </div>
        </div>
    </div>
  `,
  styles: [`
    ::ng-deep .chat-link { color: #3b82f6; text-decoration: none; word-break: break-all; }
    ::ng-deep .chat-link:hover { text-decoration: underline; }
    .widget-toast {
        position: absolute; bottom: 70px; left: 50%; transform: translateX(-50%);
        background: rgba(0,0,0,0.8); color: white; padding: 6px 16px; border-radius: 20px;
        font-size: 13px; z-index: 10000; pointer-events: none; animation: fadeIn 0.2s ease-out; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    @keyframes fadeIn { from { opacity: 0; transform: translate(-50%, 10px); } to { opacity: 1; transform: translate(-50%, 0); } }
    
    .image-preview-overlay {
        position: fixed; top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.9); z-index: 100000;
        display: flex; justify-content: center; align-items: center;
        backdrop-filter: blur(8px);
    }
    .image-preview-header {
        position: absolute; top: 24px; right: 32px; display: flex; align-items: center; z-index: 100001;
    }
    .preview-actions {
        display: flex; align-items: center; gap: 16px;
    }
    .preview-actions .action-icon {
        color: white; font-size: 28px; cursor: pointer; transition: transform 0.2s; user-select: none;
    }
    .preview-actions .action-icon:hover { transform: scale(1.1); }
    .preview-actions .action-btn {
        width: 36px; height: 36px; border-radius: 50%; background: white; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: transform 0.2s; box-shadow: 0 4px 12px rgba(0,0,0,0.2); user-select: none;
    }
    .preview-actions .action-btn:hover { transform: scale(1.1); }
    
    .preview-nav-left, .preview-nav-right {
        position: absolute; top: 50%; transform: translateY(-50%); width: 48px; height: 48px; border-radius: 50%; background: rgba(255,255,255,0.1); display: flex; justify-content: center; align-items: center; cursor: pointer; z-index: 100001; transition: all 0.2s; user-select: none;
    }
    .preview-nav-left:hover, .preview-nav-right:hover { background: rgba(255,255,255,0.2); transform: translateY(-50%) scale(1.1); }
    .preview-nav-left span, .preview-nav-right span { color: white; font-size: 32px; }
    .preview-nav-left { left: 32px; }
    .preview-nav-right { right: 32px; }
    
    .image-preview-content {
        position: relative; max-width: 90vw; max-height: 90vh;
        display: flex; justify-content: center; align-items: center;
    }
    .preview-full-image {
        max-width: 100%; max-height: 90vh; border-radius: 4px; object-fit: contain; box-shadow: 0 4px 30px rgba(0,0,0,0.5); transition: transform 0.1s ease-out;
    }
    
    .widget-panel { 
        position: fixed; bottom: 48px; left: 16px; width: 340px; 
        border-radius: 12px; z-index: 10000; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 8px 24px rgba(0,0,0,0.15); height: 500px; 
        --bg-color: #ffffff;
        --text-primary: #202124;
        --text-secondary: #5f6368;
        --border-color: #e0e0e0;
        --hover-bg: #f1f3f4;
        --list-hover: #f8f9fa;
        --body-bg: #f4f5f7;
        --input-bg: #ffffff;
        --input-border: #dadce0;
        --msg-other-bg: #ffffff;
        --msg-other-text: #202124;
        --msg-mine-bg: linear-gradient(135deg, #007bff, #8a2be2);
        --msg-mine-text: #ffffff;
        --header-bg: linear-gradient(135deg, #007bff, #8a2be2);
        --header-text: #ffffff;
        --header-icon: rgba(255,255,255,0.8);
        --header-icon-hover: rgba(255,255,255,0.2);
        --date-badge-bg: #ffffff;
        --date-badge-text: #5f6368;
        --date-badge-border: #e0e0e0;
        --ac-bg: #ffffff;
        --ac-icon-bg: #e0e7ff;
        --ac-icon-color: #2563eb;
        background: var(--bg-color);
    }
    
    @media (prefers-color-scheme: dark) {
        .widget-panel {
            --bg-color: #0f172a;
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --border-color: #1e293b;
            --hover-bg: #334155;
            --list-hover: #1e293b;
            --body-bg: #111827;
            --input-bg: #1e293b;
            --input-border: #334155;
            --msg-other-bg: #1e293b;
            --msg-other-text: #f8fafc;
            --date-badge-bg: #1e293b;
            --date-badge-text: #94a3b8;
            --date-badge-border: #334155;
            --ac-bg: #1e293b;
            --ac-icon-bg: #334155;
            --ac-icon-color: #60a5fa;
            box-shadow: 0 8px 24px rgba(0,0,0,0.5);
        }
    }

    .wp-header-white { padding: 16px; display: flex; justify-content: space-between; align-items: center; border-top-left-radius: 12px; border-top-right-radius: 12px; }
    .wp-header-chat { background: var(--header-bg); padding: 12px 16px; display: flex; justify-content: space-between; align-items: flex-start; }
    .wp-header-chat .icon-btn { color: var(--header-icon); cursor: pointer; font-size: 20px; transition: 0.2s; border-radius: 4px; padding: 4px; }
    .wp-header-chat .icon-btn:hover { background: var(--header-icon-hover); color: var(--header-text); }
    .wp-header-chat .icon-btn.active { background: rgba(255,255,255,0.3); color: #fff; }
    
    .icon-btn { color: var(--text-secondary); cursor: pointer; font-size: 20px; transition: 0.2s; border-radius: 4px; padding: 4px;}
    .msg-hover-actions span:hover { background: var(--hover-bg); border-radius: 4px; }
    
    .msg-reaction-popover {
        position: absolute; top: -35px; background: var(--bg-color);
        border: 1px solid var(--border-color); border-radius: 20px;
        padding: 4px 8px; display: flex; gap: 6px; z-index: 10;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1); align-items: center;
    }
    .msg-row.mine .msg-reaction-popover { right: 0; }
    .msg-row.other .msg-reaction-popover { left: 0; }
    .msg-reaction-popover span { cursor: pointer; font-size: 16px; transition: transform 0.1s; display: inline-flex; align-items: center;}
    .msg-reaction-popover span:hover { transform: scale(1.2); }
    .msg-reactions-display {
        position: absolute; bottom: -12px; right: 8px; background: var(--bg-color);
        border: 1px solid var(--border-color); border-radius: 12px;
        padding: 2px 6px; display: flex; gap: 4px; z-index: 5; font-size: 11px;
    }
    .msg-reactions-display span { cursor: pointer; }
    
    .chat-close { color: rgba(255,255,255,0.7); cursor: pointer; border-radius: 50%; padding: 4px; transition: 0.2s; background: rgba(255,255,255,0.1); font-size: 18px; }
    .chat-close:hover { background: rgba(255,255,255,0.2); color: #fff; }
    
    .chat-av { width: 40px; height: 40px; border-radius: 50%; background: #e0e0e0; border: 2px solid #fff; overflow: hidden; display: flex; align-items: center; justify-content: center; }
    .chat-av img { width: 100%; height: 100%; object-fit: cover; }
    
    .wp-search { padding: 12px 16px; position: relative; border-bottom: 1px solid var(--border-color); background: var(--bg-color); }
    .wp-search input { width: 100%; border: none; outline: none; font-size: 13px; color: var(--text-primary); padding-left: 28px; background: transparent; }
    .wp-search input::placeholder { color: var(--text-secondary); }
    .wp-search .material-symbols-outlined { position: absolute; left: 16px; top: 11px; font-size: 18px; color: var(--text-secondary); }

    .wp-section-title { font-size: 11px; font-weight: 700; color: var(--text-secondary); padding: 12px 16px 8px; letter-spacing: 0.5px; background: var(--bg-color); }
    
    .wp-body-list { flex: 1; overflow-y: auto; padding-bottom: 8px; background: var(--bg-color); }
    .chat-list-item { display: flex; align-items: center; padding: 10px 16px; cursor: pointer; transition: background 0.2s; }
    .chat-list-item:hover { background: var(--list-hover); }
    
    .cli-av { width: 40px; height: 40px; border-radius: 50%; color: #fff; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: 600; margin-right: 12px; position: relative; flex-shrink: 0; }
    .cli-av img { width: 100%; height: 100%; object-fit: cover; border-radius: 50%; }
    .cli-status { position: absolute; bottom: 0; right: 0; width: 12px; height: 12px; border-radius: 50%; border: 2px solid var(--bg-color); }
    .cli-status.online { background: #22c55e; }
    .cli-status.offline { background: #9ca3af; }
    
    .cli-info { flex: 1; min-width: 0; border-bottom: 1px solid var(--border-color); padding-bottom: 10px; margin-bottom: -10px; }
    .cli-top { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 2px; }
    .cli-name { font-size: 14px; font-weight: 600; color: var(--text-primary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .cli-time { font-size: 11px; color: var(--text-secondary); flex-shrink: 0; margin-left: 8px; }
    .cli-preview { font-size: 13px; color: var(--text-secondary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    
    .chat-body-area { position: relative; padding: 12px; display: flex; flex-direction: column; gap: 8px; flex: 1; background: var(--body-bg); overflow-y: auto; }
    
    .date-badge { text-align: center; margin: 12px 0; }
    .date-badge span { background: var(--date-badge-bg); color: var(--date-badge-text); font-size: 11px; font-weight: 600; padding: 4px 12px; border-radius: 12px; border: 1px solid var(--date-badge-border); display: inline-block; box-shadow: 0 1px 2px rgba(0,0,0,0.05); transition: 0.2s; }
    .hover-date-badge:hover { background: #1a73e8 !important; color: white !important; border-color: #1a73e8 !important; }

    @keyframes badgeHighlightPulse {
        0% { transform: scale(1); box-shadow: 0 0 0 rgba(26,115,232,0); background: var(--date-badge-bg); color: var(--date-badge-text); border-color: var(--date-badge-border); }
        15% { transform: scale(1.15); box-shadow: 0 0 20px rgba(26,115,232,0.7); background: #1a73e8; color: white; border-color: #1a73e8; }
        85% { transform: scale(1.15); box-shadow: 0 0 20px rgba(26,115,232,0.7); background: #1a73e8; color: white; border-color: #1a73e8; }
        100% { transform: scale(1); box-shadow: 0 0 0 rgba(26,115,232,0); background: var(--date-badge-bg); color: var(--date-badge-text); border-color: var(--date-badge-border); }
    }
    .badge-pulse { animation: badgeHighlightPulse 1.5s ease-in-out !important; }

    /* CUSTOM CALENDAR STYLES */
    .calendar-modal-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.3); display: flex; align-items: center; justify-content: center; z-index: 1000; padding: 16px; }
    .calendar-modal { background: #1f2937; border-radius: 16px; padding: 20px; width: 100%; max-width: 320px; box-shadow: 0 10px 25px rgba(0,0,0,0.5); font-family: 'Inter', sans-serif; }
    .cal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
    .cal-nav-btn { background: rgba(255,255,255,0.08); border: none; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #cbd5e1; cursor: pointer; transition: 0.2s; }
    .cal-nav-btn:hover { background: rgba(255,255,255,0.15); color: #fff; }
    .cal-month-title { color: #f8fafc; font-weight: 700; font-size: 16px; letter-spacing: 0.5px; }
    .cal-weekdays { display: grid; grid-template-columns: repeat(7, 1fr); text-align: center; color: #94a3b8; font-size: 12px; font-weight: 600; margin-bottom: 12px; }
    .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; text-align: center; }
    .cal-cell { position: relative; height: 36px; display: flex; align-items: center; justify-content: center; font-size: 14px; color: #f1f5f9; font-weight: 600; border-radius: 8px; cursor: pointer; transition: 0.2s; }
    .cal-cell:not(.empty):hover { background: rgba(255,255,255,0.1); }
    .cal-cell.empty { cursor: default; }
    .msg-bubble { padding: 8px 12px; border-radius: 12px; font-size: 13px; line-height: 1.4; overflow-wrap: anywhere; word-break: break-word; position: relative; }
    .msg-row { display: flex; align-items: flex-end; max-width: 85%; position: relative; margin-bottom: 4px; }
    .msg-row.mine { align-self: flex-end; }
    .msg-row.other { align-self: flex-start; }
    .msg-row.mine .msg-bubble { background: var(--msg-mine-bg); color: var(--msg-mine-text); border-bottom-right-radius: 4px; }
    .msg-row.other .msg-bubble { background: var(--msg-other-bg); color: var(--msg-other-text); border: 1px solid var(--border-color); border-bottom-left-radius: 4px; }
    .msg-row.mine .msg-bubble.emoji-only, .msg-row.other .msg-bubble.emoji-only { background: transparent !important; border: none !important; box-shadow: none !important; font-size: 32px; padding: 0; }
    
    .msg-file-card { display: flex; align-items: center; gap: 12px; text-decoration: none; padding: 8px 12px; border-radius: 12px; max-width: 240px; position: relative; }
    .msg-row.mine .msg-file-card { background: var(--msg-mine-bg); color: var(--msg-mine-text); border-bottom-right-radius: 4px; }
    .msg-row.other .msg-file-card { background: var(--ac-bg); color: var(--text-primary); box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid var(--border-color); border-bottom-left-radius: 4px; }
    
    .msg-image-wrapper { position: relative; max-width: 240px; border-radius: 12px; display: flex; }
    .msg-row.mine .msg-image-wrapper { border-bottom-right-radius: 4px; }
    .msg-row.other .msg-image-wrapper { border-bottom-left-radius: 4px; border: 1px solid var(--border-color); box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .chat-image-preview { width: 100%; height: auto; max-height: 240px; object-fit: cover; display: block; cursor: pointer; transition: transform 0.2s; border-radius: inherit; }
    .chat-image-preview:hover { transform: scale(1.02); }
    
    .ac-icon { width: 32px; height: 32px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .msg-row.mine .ac-icon { background: rgba(255,255,255,0.2); color: #fff; }
    .msg-row.other .ac-icon { background: var(--ac-icon-bg); color: var(--ac-icon-color); }
    .ac-name { font-weight: 500; font-size: 13px; overflow-wrap: anywhere; word-break: break-all; white-space: normal; line-height: 1.2; }

    .msg-meta { font-size: 10px; color: var(--text-secondary); margin-top: 4px; display: flex; align-items: center; }
    .msg-row.mine .msg-meta { justify-content: flex-end; }
    .msg-row.other .msg-meta { justify-content: flex-start; }
    
    .msg-row .msg-hover-actions { display:none; position:absolute; top:-16px; background:var(--bg-color); color:var(--text-secondary); border-radius:12px; padding:4px 8px; gap:6px; box-shadow:0 2px 5px rgba(0,0,0,0.15); border: 1px solid var(--border-color); z-index:10; align-items:center; white-space:nowrap; }
    .msg-row.mine .msg-hover-actions { right: 0; }
    .msg-row.other .msg-hover-actions { left: 0; }
    .msg-row:hover .msg-hover-actions { display:flex; }
    .msg-hover-actions span { font-size:16px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; }
    .msg-hover-actions span:hover { color:var(--text-primary); }

    .chat-input-area { padding: 8px 12px; border-top: 1px solid var(--border-color); display: flex; align-items: flex-end; gap: 8px; background: var(--bg-color); position: relative; }
    .chat-input-area .icon-btn { color: var(--text-secondary); cursor: pointer; font-size: 24px; transition: 0.2s; border-radius: 50%; padding: 4px; margin-bottom: 2px; }
    .chat-input-area .icon-btn:hover { background: var(--hover-bg); }
    .chat-input-area .chat-input-textarea { flex: 1; background: var(--input-bg); border: 1px solid var(--input-border); color: var(--text-primary); border-radius: 20px; padding: 10px 14px; outline: none; font-size: 13px; transition: border-color 0.2s; resize: none; overflow-y: auto; max-height: 120px; min-height: 40px; box-sizing: border-box; font-family: inherit; line-height: 20px; word-break: break-word; scrollbar-width: none; }
    .chat-input-area .chat-input-textarea::-webkit-scrollbar { display: none; }
    .chat-input-area .chat-input-textarea:focus { border-color: #8a2be2; }
    .chat-input-area .chat-input-textarea::placeholder { color: var(--text-secondary); }
    
    .send-btn { width: 36px; height: 36px; border-radius: 50%; background: linear-gradient(135deg, #007bff, #8a2be2); color: #fff; display: flex; align-items: center; justify-content: center; cursor: pointer; border: none; outline: none; box-shadow: 0 2px 4px rgba(0,0,0,0.2); flex-shrink: 0; margin-bottom: 2px; }
    .send-btn:hover { opacity: 0.9; transform: scale(1.05); transition: 0.2s; }
    
    .media-picker-inline { display: block; height: 320px; background: var(--bg-color); }

    .chat-input-container { display: flex; flex-direction: column; background: var(--bg-color); border-top: 1px solid var(--border-color); }
    .staged-file { display: inline-flex; align-items: center; gap: 8px; background: rgba(138, 43, 226, 0.1); border: 1px solid rgba(138, 43, 226, 0.5); padding: 6px 12px; border-radius: 16px; margin: 8px 12px 0 12px; align-self: flex-start; color: var(--text-primary); font-size: 13px; }
    .staged-file .close-btn { font-size: 16px; cursor: pointer; color: var(--text-secondary); transition: color 0.2s; }
    .staged-file .close-btn:hover { color: #ef4444; }
    .staged-name { max-width: 200px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

    .selection-toolbar { display: flex; align-items: center; justify-content: space-between; padding: 10px 16px; background: var(--bg-color); border-bottom: 1px solid var(--border-color); }
    .st-btn { padding: 4px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; border: none; outline: none; transition: 0.2s; }
    .st-all { background: var(--hover-bg); color: var(--text-primary); border: 1px solid var(--border-color); }
    .st-all:hover { background: var(--border-color); }
    .st-cancel { background: var(--hover-bg); color: var(--text-primary); border: 1px solid var(--border-color); }
    .st-cancel:hover { background: var(--border-color); }
    .st-delete { background: #ef4444; color: #fff; }
    .st-delete:hover { background: #dc2626; }
    .st-forward { background: #3b82f6; color: #fff; }
    .st-forward:hover { background: #2563eb; }

    .upload-spinner {
        border: 2px solid rgba(255,255,255,0.3);
        border-radius: 50%;
        border-top: 2px solid #fff;
        width: 16px;
        height: 16px;
        animation: spin 1s linear infinite;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    .reply-banner {
        display: flex;
        align-items: center;
        background: rgba(59, 130, 246, 0.1);
        padding: 8px 12px;
        margin-bottom: 8px;
        border-radius: 8px;
        border-left: 4px solid #3b82f6;
    }
    .reply-content {
        flex: 1;
        display: flex;
        flex-direction: column;
        font-size: 12px;
    }
    .reply-content strong { color: #3b82f6; }
    .reply-content span { color: var(--text-secondary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 200px; }
    .reply-banner .material-symbols-outlined { cursor: pointer; color: var(--text-secondary); font-size: 18px; }

    .inline-reply {
        background: rgba(255, 255, 255, 0.2);
        border-left: 3px solid #fff;
        padding: 4px 8px;
        margin-bottom: 4px;
        border-radius: 4px;
        font-size: 11px;
        display: flex;
        flex-direction: column;
    }
    .other .inline-reply {
        background: rgba(0, 0, 0, 0.05);
        border-left: 3px solid #3b82f6;
    }
    .inline-reply strong { color: inherit; font-weight: bold; margin-bottom: 2px; }
    .inline-reply span { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 150px; }
    .other .inline-reply strong { color: #3b82f6; }
    .reply-banner .material-symbols-outlined:hover { color: #dc2626; }

    .forward-modal-overlay {
        position: absolute; top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.4); z-index: 100000;
        display: flex; justify-content: center; align-items: center;
    }
    .forward-modal-content {
        background: var(--bg-color); width: 90%; max-height: 90%;
        border-radius: 12px; display: flex; flex-direction: column;
        overflow: hidden; box-shadow: 0 12px 32px rgba(0,0,0,0.2);
    }
    .fm-header {
        display: flex; align-items: center; padding: 12px 16px; border-bottom: 1px solid var(--border-color);
    }
    .fm-header h3 { margin: 0; font-size: 16px; font-weight: 600; color: var(--text-primary); }
    .fm-search {
        display: flex; align-items: center; padding: 8px 16px; border-bottom: 1px solid var(--border-color);
    }
    .fm-search input {
        border: none; outline: none; background: transparent; width: 100%; margin-left: 8px; color: var(--text-primary);
    }
    .fm-list {
        flex: 1; overflow-y: auto; padding: 8px 0;
    }
    .fm-item {
        display: flex; align-items: center; padding: 8px 16px; cursor: pointer; gap: 12px;
    }
    .fm-item:hover { background: var(--hover-bg); }
    .fm-av { width: 32px; height: 32px; border-radius: 50%; display: flex; justify-content: center; align-items: center; color: #fff; font-weight: bold; font-size: 14px; }
    .fm-info { flex: 1; display: flex; flex-direction: column; }
    .fm-info strong { font-size: 14px; color: var(--text-primary); }
    .fm-info span { font-size: 12px; color: var(--text-secondary); }
    .fm-footer {
        display: flex; align-items: center; justify-content: space-between; padding: 12px 16px; border-top: 1px solid var(--border-color);
        background: var(--body-bg); font-size: 13px; color: var(--text-secondary);
    }
    .group-modal-overlay {
        position: absolute; top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.5); z-index: 100000; display: flex;
        justify-content: center; align-items: center; padding: 16px;
    }
    .group-modal {
        background: var(--bg-color); width: 100%; max-width: 360px;
        border-radius: 12px; display: flex; flex-direction: column;
        overflow: hidden; box-shadow: 0 12px 32px rgba(0,0,0,0.25); border: 1px solid var(--border-color);
    }
    .gm-header {
        display: flex; align-items: center; padding: 14px 16px; gap: 10px; border-bottom: 1px solid var(--border-color); background: var(--bg-color);
    }
    .gm-body {
        padding: 16px; background: var(--bg-color); max-height: 70vh; overflow-y: auto;
    }
    .gm-save-btn {
        width: 100%; padding: 10px; border-radius: 8px; border: none; background: #3b82f6; color: white; font-weight: 600; cursor: pointer; transition: 0.2s;
    }
    .gm-save-btn:hover { background: #2563eb; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.4); }
    .gm-delete-btn:hover { background: rgba(239,68,68,0.15) !important; border-color: rgba(239,68,68,0.4) !important; }
    .gm-search {
        display: flex; align-items: center; padding: 10px 16px; border-bottom: 1px solid var(--border-color); background: var(--bg-color);
    }
    .gm-search input {
        border: none; outline: none; background: transparent; width: 100%; margin-left: 8px; color: var(--text-primary); font-size: 13px;
    }
    .gm-search .material-symbols-outlined { color: var(--text-secondary); font-size: 18px; }
    .gm-add-btn:hover { background: rgba(59, 130, 246, 0.15) !important; }
    .rm-member-btn:hover { transform: scale(1.1); transition: 0.1s; }
  `]
})
export class ChatWidgetComponent implements OnInit, OnDestroy, OnChanges {
  @Input() activeWidget: string | null = null;
  @Output() close = new EventEmitter<void>();

  @ViewChild('scrollMe') scrollMe!: ElementRef;

  currentUser: any = null;
  conversations: any[] = [];
  channels: any[] = [];
  allUsers: any[] = [];
  serverContacts: any[] = [];
  serverGroups: any[] = [];
  
  showCreateGroup: boolean = false;
  newGroupName: string = '';
  newGroupDescription: string = '';
  newGroupAvatar: File | null = null;
  
  openedChatId: number | null = null;
  openedChatUser: any = null;
  openedChannelId: number | null = null;
  openedChannelName: string = '';
  openedChannelAvatarUrl: string | null = null;
  messages: any[] = [];
  channelMessages: any[] = [];
  
  newMessage = '';
  replyingTo: any = null;
  forwardingMessage: any = null;
  showForwardModal = false;
  forwardSearch = '';
  forwardSelection = new Set<string>();
  searchQuery = '';
  stagedFile: File | null = null;
  deletingMessageId: any = null;
  reactingMessageId: any = null;
  
  editingMessageId: any = null;
  editingMessageText: string = '';
  toastMessage: string | null = null;
  private toastTimeout: any;
  previewImageUrl: string | null = null;
  previewZoomLevel: number = 1;
  previewImagesList: any[] = [];
  previewCurrentIndex: number = -1;

  showEmojiPicker = false;
  selectionMode = false;
  selectedMessages = new Set<number>();
  emojis = ['😀', '😂', '🥰', '😎', '👍', '🙏', '🔥', '🎉', '😢', '😡', '🤔', '👀', '❤️', '✨', '💯', '🙌'];

  showGroupInfoModal = false;
  editGroupName = '';
  editGroupDesc = '';
  channelMembers: any[] = [];
  isAdmin = false;
  
  showAddMemberModal = false;
  addMemberSearchQuery = '';
  addMemberContacts: any[] = [];

  private sub: Subscription = new Subscription();

  constructor(private api: ApiService, private auth: AuthService, private chatSocket: ChatSocketService, private http: HttpClient) {}

  ngOnInit() {
    this.currentUser = this.auth.user;
    if (this.currentUser) {
        this.chatSocket.connect(this.auth.token || '');
    }

    this.pollData();

    this.sub = this.chatSocket.newMessage$.subscribe(data => {
        if (data.type === 'new_message') {
            if (this.openedChatId && (data.from_user_id === this.openedChatUser?.id || data.from_user_id === this.currentUser?.id)) {
                this.api.getChatMessages(this.openedChatId).subscribe(res => {
                    if (res && res.messages) {
                        this.messages = res.messages;
                        setTimeout(() => this.scrollToBottom(), 100);
                    }
                });
            } else {
                this.pollData();
            }
        } else if (data.type === 'channel_new_message') {
            if (this.openedChannelId !== null && this.openedChannelId === data.channel_id) {
                this.api.getChannelMessages(data.channel_id).subscribe(res => {
                    if (res && res.messages) {
                        this.channelMessages = res.messages;
                        setTimeout(() => this.scrollToBottom(), 100);
                    }
                });
            } else {
                this.pollData();
            }
        }
    });
  }

  ngOnDestroy() {
      this.sub.unsubscribe();
      this.chatSocket.disconnect();
  }

  ngOnChanges(changes: SimpleChanges) {
      if (changes['activeWidget'] && !changes['activeWidget'].firstChange) {
          this.pollData();
      }
  }

  pollData() {
    const uid = this.currentUser?.id;
    if (this.activeWidget === 'chat') {
        if (this.openedChatId) {
            this.api.getChatMessages(this.openedChatId).subscribe(res => {
                if (res && res.messages) {
                    this.messages = res.messages;
                    this.openedChatUser = res.other_user;
                }
            });
        } else {
            this.api.getChatConversations().subscribe(res => {
                this.conversations = Array.isArray(res) ? res : [];
            });
        }
    } else if (this.activeWidget === 'channels') {
        if (this.openedChannelId !== null) {
            this.api.getChannelMessages(this.openedChannelId).subscribe(res => {
                this.channelMessages = res?.messages || (Array.isArray(res) ? res : []);
            });
        } else {
            this.api.getChatChannels().subscribe(res => {
                this.channels = Array.isArray(res) ? res : [];
            });
        }
    }
  }

  filteredConversations() {
      if (!this.conversations || !Array.isArray(this.conversations)) return [];
      if (!this.searchQuery) return this.conversations;
      const q = this.searchQuery.toLowerCase();
      return this.conversations.filter(c => c?.other_user?.name && c.other_user.name.toLowerCase().includes(q));
  }

  filteredChannels() {
      if (!this.channels || !Array.isArray(this.channels)) return [];
      if (!this.searchQuery) return this.channels;
      const q = this.searchQuery.toLowerCase();
      return this.channels.filter(c => c?.name && c.name.toLowerCase().includes(q));
  }

  filteredContacts() {
      if (!this.searchQuery || !this.allUsers || !Array.isArray(this.allUsers)) return [];
      const q = this.searchQuery.toLowerCase();
      const existingConvIds = new Set(
          (this.conversations || []).map(c => c?.other_user?.id ? String(c.other_user.id) : null).filter(Boolean)
      );
      return this.allUsers.filter(u => {
          if (!u?.name) return false;
          const uid = String(u.id);
          if (existingConvIds.has(uid)) return false;
          return u.name.toLowerCase().includes(q);
      });
  }

  onSearchChange(q: string) {
      if (q && q.trim().length >= 2) {
          this.api.searchUsers(q.trim()).subscribe({
              next: (res: any) => {
                  const existingUserIds = new Set(this.conversations.map(c => String(c.other_user.id)));
                  this.serverContacts = (res || []).filter((u: any) => 
                      !existingUserIds.has(String(u.id)) && u.id !== this.currentUser?.id
                  );
              },
              error: (err: any) => {
                  console.error("Search error", err);
                  this.serverContacts = [];
              }
          });
      } else {
          this.serverContacts = [];
      }
  }

  onGroupSearchChange(q: string) {
      if (q && q.trim().length >= 1) {
          this.api.searchChannels(q.trim()).subscribe({
              next: (res: any) => {
                  this.serverGroups = res || [];
              },
              error: (err: any) => {
                  console.error("Group search error", err);
                  this.serverGroups = [];
              }
          });
      } else {
          this.serverGroups = [];
      }
  }

  resetCreateGroup() {
      this.newGroupName = '';
      this.newGroupDescription = '';
      this.newGroupAvatar = null;
  }

  onGroupAvatarSelected(event: any) {
      const file = event.target.files[0];
      if (file) {
          this.newGroupAvatar = file;
      }
      event.target.value = '';
  }

  createGroup() {
      const name = this.newGroupName?.trim();
      if (!name) return;
      
      this.api.createChannel(name, this.newGroupDescription).subscribe({
          next: (res: any) => {
              this.channels.unshift(res);
              this.showCreateGroup = false;
              this.resetCreateGroup();
              this.openChannel(res.id, res.name, res.avatar_url);
          },
          error: (err: any) => {
              console.error("Failed to create group", err);
          }
      });
  }

  filteredServerGroups() {
      if (!this.searchQuery || !this.serverGroups || !Array.isArray(this.serverGroups)) return [];
      const existingChannelIds = new Set((this.channels || []).map(c => String(c.id)));
      return this.serverGroups.filter(c => !existingChannelIds.has(String(c.id)));
  }

  openWidgetChat(userId: number) {
      this.activeWidget = 'chat';
      this.openChat(userId);
  }

  openChat(userId: number) {
      this.openedChatId = userId;
      this.messages = [];
      this.pollData();
  }

  openChannel(channelId: number, channelName: string, avatarUrl: string | null = null) {
      this.openedChannelId = channelId;
      this.openedChannelName = channelName;
      this.openedChannelAvatarUrl = avatarUrl;
      this.channelMessages = [];
      this.loadChannelInfo(channelId);
      this.pollData();
  }

  loadChannelInfo(channelId: number) {
      const channel = this.channels.find(c => c.id === channelId);
      if (channel) {
          this.editGroupName = channel.name;
          this.editGroupDesc = channel.description;
      }
      this.api.getChannelMembers(channelId).subscribe({
          next: (members) => {
              this.channelMembers = members || [];
              const myMember = this.channelMembers.find(m => m.id === this.currentUser?.id);
              this.isAdmin = myMember ? (myMember.role === 'admin' || myMember.is_admin === 1 || myMember.is_admin === true) : false;
          },
          error: (err) => {
              console.error("Failed to load channel members", err);
              this.channelMembers = [];
          }
      });
  }

  openGroupInfo() {
      if (!this.openedChannelId) return;
      this.loadChannelInfo(this.openedChannelId);
      this.showGroupInfoModal = true;
  }

  saveGroupInfo() {
      if (!this.openedChannelId) return;
      this.api.updateChannel(this.openedChannelId, this.editGroupName, this.editGroupDesc).subscribe(res => {
          this.openedChannelName = this.editGroupName;
          this.showGroupInfoModal = false;
          this.pollData(); // Refresh list if needed
      });
  }

  deleteGroup() {
      if (!this.openedChannelId) return;
      if (!confirm("Are you sure you want to delete this group for everyone?")) return;
      this.api.deleteChannel(this.openedChannelId).subscribe(() => {
          this.showGroupInfoModal = false;
          this.closeChannel();
          this.pollData();
      });
  }

  leaveGroup() {
      if (!this.openedChannelId) return;
      if (!confirm("Are you sure you want to leave this group?")) return;
      this.api.leaveChannel(this.openedChannelId).subscribe(() => {
          this.closeChannel();
          this.pollData();
      });
  }

  removeMember(userId: number) {
      if (!this.openedChannelId) return;
      if (!confirm("Remove this member?")) return;
      this.api.removeChannelMember(this.openedChannelId, userId).subscribe(() => {
          this.loadChannelInfo(this.openedChannelId!);
      });
  }

  setMemberAdmin(userId: number, isAdmin: boolean) {
      if (!this.openedChannelId) return;
      this.api.setChannelAdmin(this.openedChannelId, userId, isAdmin).subscribe(() => {
          this.loadChannelInfo(this.openedChannelId!);
      });
  }

  openAddMember() {
      if (!this.openedChannelId) return;
      this.showAddMemberModal = true;
      this.addMemberSearchQuery = '';
      this.addMemberContacts = this.allUsers; // Using all users or contacts for simplicity
  }

  onAddMemberSearchChange() {
      // filtering handled in getter
  }

  filteredAddMemberContacts() {
      const currentMemberIds = new Set(this.channelMembers.map(m => m.id));
      return this.addMemberContacts.filter(c => 
          !currentMemberIds.has(c.id) &&
          c.name.toLowerCase().includes(this.addMemberSearchQuery.toLowerCase())
      );
  }

  addMember(userId: number) {
      if (!this.openedChannelId) return;
      this.api.addChannelMembers(this.openedChannelId, [userId]).subscribe(() => {
          this.addMemberSearchQuery = '';
          this.loadChannelInfo(this.openedChannelId!);
          this.showAddMemberModal = false;
      });
  }

  closeChannel() {
      this.openedChannelId = null;
      this.pollData();
  }

  closeChat() {
      this.openedChatId = null;
      this.pollData();
  }

  resetInputHeight() {
      setTimeout(() => {
          const textareas = document.querySelectorAll('.chat-input-textarea') as NodeListOf<HTMLTextAreaElement>;
          textareas.forEach(ta => ta.style.height = 'auto');
      }, 0);
  }

  onInputKeydown(event: KeyboardEvent) {
      if (event.key === 'Enter' && !event.shiftKey) {
          event.preventDefault();
          this.sendMessage();
      }
  }

  autoResizeInput(event: Event) {
      const target = event.target as HTMLTextAreaElement;
      if (target) {
          target.style.height = 'auto';
          target.style.height = Math.min(target.scrollHeight, 120) + 'px';
      }
  }


  sendMedia(media: {url: string, type: string}) {
      const prefix = media.type === 'gif' ? '[GIF]' : '[Sticker]';
      const finalMsg = `${prefix}${media.url}`;
  
      this.showEmojiPicker = false;
      
      if (this.activeWidget === 'chat' && this.openedChatId) {
          this.chatSocket.forceStopTyping(this.openedChatId);
          this.api.sendChatMessage(this.openedChatId, finalMsg).subscribe(res => {
              this.messages.push(res);
              setTimeout(() => this.scrollToBottom(), 100);
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.chatSocket.forceStopChannelTyping(this.openedChannelId);
          this.api.sendChannelMessage(this.openedChannelId, finalMsg).subscribe(res => {
              this.channelMessages.push(res);
              setTimeout(() => this.scrollToBottom(), 100);
          });
      }
  }

  sendMessage() {
      const text = this.newMessage.trim();
      const file = this.stagedFile;
      if (!text && !file) return;

      let finalMsg = text;
      if (this.replyingTo && text) {
          finalMsg = `[REPLY_META:${this.replyingTo.id}]\n${text}`;
          this.replyingTo = null;
      }

      this.newMessage = '';
      this.stagedFile = null;
      this.showEmojiPicker = false;
      this.resetInputHeight();
      
      if (this.activeWidget === 'chat' && this.openedChatId) {
          this.chatSocket.forceStopTyping(this.openedChatId);
          if (text) {
              this.api.sendChatMessage(this.openedChatId, finalMsg).subscribe(res => {
                  this.messages.push(res);
                  setTimeout(() => this.scrollToBottom(), 100);
              });
          }
          if (file) {
              const tempId = 'temp-' + Date.now();
              const tempMsg = {
                  id: tempId,
                  is_mine: true,
                  is_file: true,
                  file_path: '',
                  message: file.name,
                  created_at: new Date().toISOString(),
                  isUploading: true
              };
              this.messages.push(tempMsg);
              setTimeout(() => this.scrollToBottom(), 100);

              const url = `${environment.apiUrl}/chat/upload-stream?to_user_id=${this.openedChatId}&filename=${encodeURIComponent(file.name)}&mime_type=${encodeURIComponent(file.type)}`;
              this.http.post(url, file).subscribe({
                  next: (msg: any) => {
                      const idx = this.messages.findIndex(m => m.id === tempId);
                      if (idx !== -1) this.messages[idx] = msg;
                      else this.messages.push(msg);
                      setTimeout(() => this.scrollToBottom(), 100);
                  },
                  error: (err) => {
                      console.error("Upload failed", err);
                      this.messages = this.messages.filter(m => m.id !== tempId);
                  }
              });
          }
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.chatSocket.forceStopChannelTyping(this.openedChannelId);
          if (text) {
              this.api.sendChannelMessage(this.openedChannelId, finalMsg).subscribe(res => {
                  this.channelMessages.push(res);
                  setTimeout(() => this.scrollToBottom(), 100);
              });
          }
          if (file) {
              const tempId = 'temp-' + Date.now();
              const tempMsg = {
                  id: tempId,
                  is_mine: true,
                  is_file: true,
                  file_path: '',
                  message: file.name,
                  created_at: new Date().toISOString(),
                  isUploading: true
              };
              this.channelMessages.push(tempMsg);
              setTimeout(() => this.scrollToBottom(), 100);

              const url = `${environment.apiUrl}/chat/channels/${this.openedChannelId}/upload-stream?filename=${encodeURIComponent(file.name)}&mime_type=${encodeURIComponent(file.type)}`;
              this.http.post(url, file).subscribe({
                  next: (msg: any) => {
                      const idx = this.channelMessages.findIndex(m => m.id === tempId);
                      if (idx !== -1) this.channelMessages[idx] = msg;
                      else this.channelMessages.push(msg);
                      setTimeout(() => this.scrollToBottom(), 100);
                  },
                  error: (err) => {
                      console.error("Upload failed", err);
                      this.channelMessages = this.channelMessages.filter(m => m.id !== tempId);
                  }
              });
          }
      }
  }

  getInitials(name: string): string {
      if (!name) return '??';
      return name.split(' ').map(n => n[0]).join('').toUpperCase().substring(0, 2);
  }

  isOnline(lastLogin: string | null | undefined): boolean {
      if (!lastLogin) return false;
      const diff = (Date.now() - new Date(lastLogin).getTime()) / 60000;
      return diff <= 10;
  }

  formatTime(isoString: string | null | undefined): string {
      if (!isoString) return '';
      const d = new Date(isoString);
      return isNaN(d.getTime()) ? '' : d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', hour12: true });
  }

  isFirstOfDay(index: number, arr: any[]): boolean {
      if (index === 0) return true;
      const d1 = new Date(arr[index].created_at);
      const d2 = new Date(arr[index - 1].created_at);
      return isNaN(d1.getTime()) || isNaN(d2.getTime()) ? false : d1.toDateString() !== d2.toDateString();
  }

  getDateLabel(isoString: string | null | undefined): string {
      if (!isoString) return '';
      const d = new Date(isoString);
      if (isNaN(d.getTime())) return '';
      const today = new Date();
      if (d.toDateString() === today.toDateString()) return 'TODAY';
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);
      if (d.toDateString() === yesterday.toDateString()) return 'YESTERDAY';
      return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  cleanMessage(text: string | null | undefined): string {
      if (!text) return '';
      return text.replace(/\[REPLY_META:[\s\S]*?\]\n?/g, '').replace(/\[Reply:[\s\S]*?\]\n?/g, '');
  }

  linkify(text: string): string {
      if (!text) return '';
      // Escape HTML to prevent XSS (since we are using innerHTML)
      let escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      // Replace URLs with anchor tags
      const urlRegex = /(https?:\/\/[^\s]+)/g;
      return escaped.replace(urlRegex, (url) => {
          return `<a href="${url}" target="_blank" class="chat-link">${url}</a>`;
      }).replace(/\n/g, '<br>');
  }

  isOnlyEmoji(text: string | null | undefined): boolean {
      if (!text) return false;
      const clean = this.cleanMessage(text).trim();
      if (!clean) return false;
      // Regex to match string containing strictly emoji characters and optional spaces
      const emojiRegex = /^[\p{Extended_Pictographic}\s]+$/u;
      return emojiRegex.test(clean);
  }

  getReplyMeta(text: string | null | undefined): any | null {
      if (!text) return null;
      const match = text.match(/\[REPLY_META:(.*?)\]/);
      if (match && match[1]) {
          const id = match[1];
          return this.messages.find(m => String(m.id) === String(id)) || this.channelMessages.find(m => String(m.id) === String(id)) || null;
      }
      return null;
  }

  isSelf(id: any): boolean {
      return this.currentUser?.id && String(id) == String(this.currentUser.id);
  }

  scrollToBottom() {
      if (this.scrollMe) this.scrollMe.nativeElement.scrollTop = this.scrollMe.nativeElement.scrollHeight;
  }

  toggleSelectionMode() {
      this.selectionMode = !this.selectionMode;
      this.selectedMessages.clear();
  }

  // --- CUSTOM CALENDAR LOGIC ---
  showCalendarModal = false;
  calendarCurrentMonth = new Date();
  calendarDays: any[] = [];
  calendarTargetContext: 'chat' | 'channel' = 'chat';
  msgDatesSet = new Set<string>();

  openCustomCalendar(context: 'chat' | 'channel') {
      this.calendarTargetContext = context;
      this.calendarCurrentMonth = new Date();
      this.generateCalendarDays();
      this.showCalendarModal = true;
  }

  changeMonth(delta: number) {
      this.calendarCurrentMonth = new Date(this.calendarCurrentMonth.getFullYear(), this.calendarCurrentMonth.getMonth() + delta, 1);
      this.generateCalendarDays();
  }

  generateCalendarDays() {
      this.calendarDays = [];
      const year = this.calendarCurrentMonth.getFullYear();
      const month = this.calendarCurrentMonth.getMonth();
      const firstDay = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();

      for (let i = 0; i < firstDay; i++) {
          this.calendarDays.push(null);
      }
      for (let i = 1; i <= daysInMonth; i++) {
          this.calendarDays.push(new Date(year, month, i));
      }

      this.msgDatesSet.clear();
      const msgList = this.calendarTargetContext === 'chat' ? this.messages : this.channelMessages;
      msgList.forEach(m => {
          const d = new Date(m.created_at);
          this.msgDatesSet.add(`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`);
      });
  }

  hasMessagesOnDate(d: Date): boolean {
      return this.msgDatesSet.has(`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`);
  }

  selectCustomDate(d: Date) {
      this.showCalendarModal = false;
      const msgList = this.calendarTargetContext === 'chat' ? this.messages : this.channelMessages;
      let targetIndex = null;
      d.setHours(0,0,0,0);
      for (let i = 0; i < msgList.length; i++) {
          const msgDate = new Date(msgList[i].created_at);
          msgDate.setHours(0,0,0,0);
          if (msgDate >= d) {
              targetIndex = i;
              break;
          }
      }
      if (targetIndex !== null) {
          setTimeout(() => {
              const badgeId = 'date-badge-' + this.calendarTargetContext + '-' + targetIndex;
              const el = document.getElementById(badgeId);
              if (el && this.scrollMe) {
                  const container = this.scrollMe.nativeElement;
                  container.scrollTop = el.offsetTop - 12;

                  const span = el.querySelector('span.hover-date-badge');
                  if (span) {
                      span.classList.remove('badge-pulse');
                      void span.clientWidth; // trigger reflow
                      span.classList.add('badge-pulse');
                      setTimeout(() => span.classList.remove('badge-pulse'), 1500);
                  }
              }
          }, 100);
      }
  }
  // -----------------------------

  scrollToDate(event: Event, msgList: any[]) {
      const target = event.target as HTMLInputElement;
      if (!target.value) return;
      
      const selectedDate = new Date(target.value);
      selectedDate.setHours(0,0,0,0);

      let targetMsgId = null;
      // find the first message that matches this date, or the closest one AFTER it
      for (let i = 0; i < msgList.length; i++) {
          const msgDate = new Date(msgList[i].created_at);
          msgDate.setHours(0,0,0,0);
          
          if (msgDate >= selectedDate) {
              targetMsgId = msgList[i].id;
              break;
          }
      }

      if (targetMsgId) {
          setTimeout(() => {
              const el = document.getElementById('msg-' + targetMsgId);
              if (el && this.scrollMe) {
                  const container = this.scrollMe.nativeElement;
                  container.scrollTop = el.offsetTop - 20; // 20px padding
              }
          }, 100);
      }
      target.value = '';
  }

  toggleMessageSelection(msg: any) {
      if (this.selectedMessages.has(msg.id)) this.selectedMessages.delete(msg.id);
      else this.selectedMessages.add(msg.id);
  }

  getFileUrl(path: string): string {
      if (!path) return '#';
      if (path.includes('/chat/download?path=')) return path;
      return `${environment.apiUrl}/chat/download?path=${encodeURIComponent(path)}`;
  }

  openFile(path: string, originalName?: string, msg?: any, isChannel?: boolean) {
      if (msg && (this.isImage(path) || this.isImage(originalName))) {
          this.openImagePreview(msg, !!isChannel);
          return;
      }
      const url = this.getFileUrl(path);
      if (url && url !== '#') {
          window.open(url, '_blank');
      }
  }

  openImagePreview(msg: any, isChannel: boolean) {
      const list = isChannel ? this.channelMessages : this.messages;
      this.previewImagesList = list.filter(m => m.is_file && (this.isImage(m.file_path) || this.isImage(m.message)));
      this.previewCurrentIndex = this.previewImagesList.findIndex(m => m.id === msg.id);
      
      this.previewZoomLevel = 1;
      this.previewImageUrl = this.getFileUrl(msg.file_path);
  }

  zoomPreview(delta: number) {
      this.previewZoomLevel = Math.max(0.2, Math.min(5, this.previewZoomLevel + delta));
  }

  onPreviewWheel(event: WheelEvent) {
      event.preventDefault();
      const delta = event.deltaY < 0 ? 0.1 : -0.1;
      this.zoomPreview(delta);
  }

  hasPrevImage(): boolean {
      return this.previewCurrentIndex > 0;
  }

  hasNextImage(): boolean {
      return this.previewCurrentIndex >= 0 && this.previewCurrentIndex < this.previewImagesList.length - 1;
  }

  prevImage() {
      if (this.hasPrevImage()) {
          this.previewCurrentIndex--;
          const msg = this.previewImagesList[this.previewCurrentIndex];
          this.previewImageUrl = this.getFileUrl(msg.file_path);
          this.previewZoomLevel = 1;
      }
  }

  nextImage() {
      if (this.hasNextImage()) {
          this.previewCurrentIndex++;
          const msg = this.previewImagesList[this.previewCurrentIndex];
          this.previewImageUrl = this.getFileUrl(msg.file_path);
          this.previewZoomLevel = 1;
      }
  }

  downloadPreviewImage() {
      if (this.previewImagesList && this.previewCurrentIndex >= 0) {
          const msg = this.previewImagesList[this.previewCurrentIndex];
          this.downloadMessage(msg);
      }
  }

  closePreview() {
      this.previewImageUrl = null;
      this.previewZoomLevel = 1;
  }

  isImage(path: string | undefined): boolean {
      if (!path) return false;
      const ext = path.split('.').pop()?.toLowerCase();
      return !!ext && ['jpg', 'jpeg', 'png', 'gif', 'webp'].includes(ext);
  }

  downloadMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      let url = this.getFileUrl(msg.file_path);
      if (msg.message) {
          url += `&filename=${encodeURIComponent(msg.message)}`;
      }
      const a = document.createElement('a');
      a.href = url;
      a.download = msg.message || 'download';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
  }

  replyToMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.replyingTo = msg;
  }

  cancelReply() {
      this.replyingTo = null;
  }

  forwardMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.forwardingMessage = msg;
      this.showForwardModal = true;
      this.forwardSearch = '';
      this.forwardSelection.clear();
      if (!this.allUsers.length) {
          this.api.getAllUsers().subscribe(res => this.allUsers = res || []);
      }
      if (!this.channels.length) {
          this.api.getChatChannels().subscribe(res => this.channels = res || []);
      }
  }

  closeForwardModal() {
      this.showForwardModal = false;
      this.forwardingMessage = null;
  }

  filteredForwardUsers(): any[] {
      const q = this.forwardSearch.toLowerCase();
      return this.allUsers.filter(u => u.name && u.name.toLowerCase().includes(q));
  }

  filteredForwardChannels(): any[] {
      const q = this.forwardSearch.toLowerCase();
      return this.channels.filter(c => c.name && c.name.toLowerCase().includes(q));
  }

  toggleForwardSelection(id: string) {
      if (this.forwardSelection.has(id)) this.forwardSelection.delete(id);
      else this.forwardSelection.add(id);
  }

  sendForward() {
      if (!this.forwardingMessage || this.forwardSelection.size === 0) return;
      
      const msgText = this.cleanMessage(this.forwardingMessage.message).trim();
      const isFile = !!this.forwardingMessage.is_file;
      const filePath = this.forwardingMessage.file_path;

      this.forwardSelection.forEach(id => {
          if (id.startsWith('user_')) {
              const uId = parseInt(id.replace('user_', ''), 10);
              this.api.sendChatMessage(uId, msgText, isFile, filePath).subscribe();
          } else if (id.startsWith('channel_')) {
              const cId = parseInt(id.replace('channel_', ''), 10);
              this.api.sendChannelMessage(cId, msgText, isFile, filePath).subscribe();
          }
      });

      this.closeForwardModal();
  }

  reactToMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.reactingMessageId = msg.id;
  }

  addReaction(msg: any, emoji: string) {
      if (this.activeWidget === 'chat') {
          this.api.reactToChatMessage(msg.id, emoji).subscribe({
              next: (res) => {
                  msg.reactions = res.reactions;
                  this.reactingMessageId = null;
              },
              error: err => console.error(err)
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.api.reactToChannelMessage(this.openedChannelId, msg.id, emoji).subscribe({
              next: (res) => {
                  msg.reactions = res.reactions;
                  this.reactingMessageId = null;
              },
              error: err => console.error(err)
          });
      }
  }

  getReactionsList(msg: any): string[] {
      if (!msg.reactions) return [];
      try {
          const dict = JSON.parse(msg.reactions);
          const counts: { [key: string]: number } = {};
          for (let uid in dict) {
              counts[dict[uid]] = (counts[dict[uid]] || 0) + 1;
          }
          return Object.keys(counts).map(k => k + (counts[k] > 1 ? ` ${counts[k]}` : ''));
      } catch { return []; }
  }

  closeReactionPopover(event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.reactingMessageId = null;
  }

  private fallbackCopy(text: string) {
      const el = document.createElement('textarea');
      el.value = text;
      el.setAttribute('readonly', '');
      el.style.position = 'absolute';
      el.style.left = '-9999px';
      document.body.appendChild(el);
      el.select();
      try {
          document.execCommand('copy');
      } catch (err) {
          console.error('Fallback: Oops, unable to copy', err);
      }
      document.body.removeChild(el);
  }

  showToast(msg: string) {
      this.toastMessage = msg;
      if (this.toastTimeout) clearTimeout(this.toastTimeout);
      this.toastTimeout = setTimeout(() => this.toastMessage = null, 2000);
  }

  copyMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      const textToCopy = this.cleanMessage(msg.message);

      if (msg.is_file && this.isImage(msg.message || msg.file_path)) {
          const url = this.getFileUrl(msg.file_path);
          fetch(url)
            .then(response => response.blob())
            .then(blob => {
              if (navigator.clipboard && window.isSecureContext && typeof ClipboardItem !== 'undefined') {
                const item = new ClipboardItem({ [blob.type]: blob });
                navigator.clipboard.write([item]).then(() => this.showToast('Image copied!')).catch(() => this.fallbackCopy(url));
              } else {
                this.fallbackCopy(url);
              }
            }).catch(() => this.fallbackCopy(url));
          return;
      }

      if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(textToCopy).then(() => this.showToast('Text copied!')).catch(() => {
              this.fallbackCopy(textToCopy);
          });
      } else {
          this.fallbackCopy(textToCopy);
          this.showToast('Text copied!');
      }
  }

  editMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      if (!msg.is_mine) {
          alert("You can only edit your own messages.");
          return;
      }
      this.editingMessageId = msg.id;
      this.editingMessageText = this.cleanMessage(msg.message);
  }

  saveEditMessage(msg: any) {
      if (!this.editingMessageText.trim()) return;
      
      const payload = { message: this.editingMessageText };
      if (this.activeWidget === 'chat') {
          this.http.put(`${this.api.base}/chat/messages/${msg.id}`, payload).subscribe({
              next: () => {
                  msg.message = this.editingMessageText;
                  msg.is_edited = 1;
                  this.editingMessageId = null;
              },
              error: err => console.error(err)
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.http.put(`${this.api.base}/chat/channels/${this.openedChannelId}/messages/${msg.id}`, payload).subscribe({
              next: () => {
                  msg.message = this.editingMessageText;
                  msg.is_edited = 1;
                  this.editingMessageId = null;
              },
              error: err => console.error(err)
          });
      }
  }

  cancelEdit() {
      this.editingMessageId = null;
  }

  copyLink(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      let link = '';
      if (msg.is_file && msg.file_path) {
          link = this.getFileUrl(msg.file_path);
      } else {
          link = window.location.origin + '/chat?msg=' + msg.id;
      }
      if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(link).then(() => this.showToast('Link copied!')).catch(() => {
              this.fallbackCopy(link);
              this.showToast('Link copied!');
          });
      } else {
          this.fallbackCopy(link);
          this.showToast('Link copied!');
      }
  }

  selectAllMessages() {
      if (this.activeWidget === 'chat') {
          this.messages.forEach(m => this.selectedMessages.add(m.id));
      } else if (this.activeWidget === 'channels') {
          this.channelMessages.forEach(m => this.selectedMessages.add(m.id));
      }
  }

  deleteSelectedMessages() {
      if (this.selectedMessages.size === 0) return;
      if (this.activeWidget === 'chat') {
          this.selectedMessages.forEach(id => {
              this.api.deleteChatMessage(id).subscribe({
                  next: () => this.messages = this.messages.filter(m => m.id !== id),
                  error: err => console.error('Failed to delete message', err)
              });
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          const cId = this.openedChannelId;
          this.selectedMessages.forEach(id => {
              this.api.deleteChannelMessage(cId, id).subscribe({
                  next: () => this.channelMessages = this.channelMessages.filter(m => m.id !== id),
                  error: err => console.error('Failed to delete channel message', err)
              });
          });
      }
      this.toggleSelectionMode();
  }

  addEmoji(emoji: string) {
      this.newMessage += emoji;
      this.showEmojiPicker = false;
  }

  onPaste(event: ClipboardEvent) {
      const items = event.clipboardData?.items;
      if (!items) return;
      for (let i = 0; i < items.length; i++) {
          if (items[i].type.indexOf('image') !== -1) {
              const file = items[i].getAsFile();
              if (file) {
                  // Give it a generic name if it doesn't have one, or keep the existing one
                  // Actually the file from clipboard usually comes as "image.png"
                  this.stagedFile = file;
                  // If we only want to paste the image and not text, uncomment below:
                  // event.preventDefault();
                  break;
              }
          }
      }
  }

  onFileSelected(event: any) {
      const file = event.target.files[0];
      if (!file) return;
      this.stagedFile = file;
      event.target.value = '';
  }

  removeStagedFile() {
      this.stagedFile = null;
  }

  deleteMessage(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.deletingMessageId = msg.id;
  }

  confirmDelete(msg: any, event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      
      if (this.activeWidget === 'chat') {
          this.api.deleteChatMessage(msg.id).subscribe({
              next: () => this.messages = this.messages.filter(m => m.id !== msg.id),
              error: err => console.error('Failed to delete message', err)
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.api.deleteChannelMessage(this.openedChannelId, msg.id).subscribe({
              next: () => this.channelMessages = this.channelMessages.filter(m => m.id !== msg.id),
              error: err => console.error('Failed to delete channel message', err)
          });
      }
      this.deletingMessageId = null;
  }

  cancelDelete(event?: Event) {
      if (event) { event.preventDefault(); event.stopPropagation(); }
      this.deletingMessageId = null;
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: Event) {
      this.showEmojiPicker = false;
  }

  onPanelClick(event: Event) {
      event.stopPropagation();
      this.showEmojiPicker = false;
  }
}
