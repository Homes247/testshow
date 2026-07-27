import { Injectable, OnDestroy } from '@angular/core';
import { HttpClient, HttpEvent, HttpEventType } from '@angular/common/http';
import { Observable, Subject } from 'rxjs';
import { map } from 'rxjs/operators';
import { environment } from '../../environments/environment';

export interface SyncMessage {
  type: 'update' | 'presence' | 'cursor' | 'cursor_remove' | 'cell_update' | 'reload_page';
  content?: string;
  title?: string;
  users?: number;
  client_id?: string;
  r?: number;
  c?: number;
  sheetIdx?: number;
  value?: any;
  formatting?: any;
  seq?: number;
}

@Injectable({ providedIn: 'root' })
export class ApiService implements OnDestroy {
  base = environment.apiUrl;
  wsBase = environment.wsUrl;
  private socket: WebSocket | null = null;
  private currentDocId: string | null = null;
  private messageSubject = new Subject<SyncMessage>();

  constructor(private http: HttpClient) { }

  createDocument(title: string, doc_type: string): Observable<any> {
    return this.http.post(`${this.base}/documents`, { title, doc_type });
  }

  getProfile(): Observable<any> {
    return this.http.get(`${this.base}/documents/profile`);
  }

  importFile(file: File, replaceDocId?: string): Observable<any> {
    const formData = new FormData();
    formData.append('file', file);
    if (replaceDocId) {
      formData.append('replace_doc_id', replaceDocId);
    }
    return this.http.post(`${this.base}/documents/import`, formData, {
      reportProgress: true,
      observe: 'events'
    });
  }

  listDocuments(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/documents`);
  }

  getDocument(id: string): Observable<any> {
    return this.http.get(`${this.base}/documents/${id}`);
  }

  saveDocument(id: string, title: string, content: string): Observable<any> {
    return this.http.put(`${this.base}/documents/${id}`, { title, content });
  }

  deleteDocument(id: string): Observable<any> {
    return this.http.delete(`${this.base}/documents/${id}`);
  }

  getAuditEvents(id: string): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/documents/${id}/audit-events`);
  }

  submitFeedback(docId: string, type: string, text: string, rating: number, recordScreen: boolean, file: File | null): Observable<any> {
    const formData = new FormData();
    formData.append('doc_id', docId);
    formData.append('type', type);
    formData.append('text', text);
    formData.append('rating', rating.toString());
    formData.append('record_screen', recordScreen ? 'true' : 'false');
    if (file) {
      formData.append('file', file);
    }
    return this.http.post(`${this.base}/documents/feedback`, formData);
  }

  saveAuditEvents(id: string, events: any[]): Observable<any> {
    return this.http.post(`${this.base}/documents/${id}/audit-events`, events);
  }

  getNotificationSettings(docId: string): Observable<any> {
    return this.http.get(`${this.base}/documents/${docId}/notification-settings`);
  }

  saveNotificationSettings(docId: string, settings: any): Observable<any> {
    return this.http.put(`${this.base}/documents/${docId}/notification-settings`, settings);
  }

  triggerNotification(docId: string, event_type: string, extra_detail?: string): Observable<any> {
    return this.http.post(`${this.base}/documents/${docId}/notify`, { event_type, extra_detail });
  }

  getSheetVersions(docId: string): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/documents/${docId}/versions`);
  }
  
  getSheetVersionSnapshot(docId: string, versionId: string): Observable<any> {
    return this.http.get<any>(`${this.base}/documents/${docId}/versions/${versionId}`);
  }

  createNamedVersion(docId: string, versionName: string): Observable<any> {
    return this.http.post<any>(`${this.base}/documents/${docId}/versions`, { version_name: versionName });
  }

  restoreSheetVersion(docId: string, versionId: string): Observable<any> {
    return this.http.post<any>(`${this.base}/documents/${docId}/versions/${versionId}/restore`, {});
  }


  restoreDocument(id: string): Observable<any> {
    return this.http.put(`${this.base}/documents/${id}`, { is_trashed: 0 });
  }

  shareDocument(docId: string, email: string, permission: string): Observable<any> {
    return this.http.post(`${this.base}/documents/${docId}/share`, { email, permission });
  }

  exportDocument(doc_id: string, format: string, password?: string): Observable<Blob> {
    const payload: any = { doc_id, format };
    if (password) payload.password = password;
    return this.http.post(`${this.base}/export`, payload, { responseType: 'blob' });
  }

  searchUsers(query: string): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/auth/search-users?q=${encodeURIComponent(query)}`);
  }

  getAllUsers(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/users`);
  }

  // CHAT ENDPOINTS
  getChatConversations(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/chat/conversations`);
  }

  getChatMessages(userId: number): Observable<any> {
    return this.http.get<any>(`${this.base}/chat/messages/${userId}`);
  }

  sendChatMessage(toUserId: number, message: string, is_file: boolean = false, file_path: string | null = null): Observable<any> {
    return this.http.post<any>(`${this.base}/chat/send`, { to_user_id: toUserId, message, is_file, file_path });
  }

  deleteChatMessage(messageId: number): Observable<any> {
    return this.http.delete<any>(`${this.base}/chat/messages/${messageId}`);
  }

  reactToChatMessage(messageId: number, emoji: string): Observable<any> {
    return this.http.put<any>(`${this.base}/chat/messages/${messageId}/react`, { emoji });
  }

  getChatChannels(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/channels/`);
  }

  searchChannels(query: string): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/channels/all`).pipe(
      map(channels => channels.filter(c => c.name.toLowerCase().includes(query.toLowerCase())))
    );
  }

  getChannelMessages(channelId: number): Observable<any> {
    return this.http.get<any>(`${this.base}/channels/${channelId}/messages`);
  }

  getChannelInfo(channelId: number): Observable<any> {
    return this.http.get<any>(`${this.base}/channels/${channelId}/info`);
  }

  getChannelMembers(channelId: number): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/channels/${channelId}/members`);
  }

  updateChannel(channelId: number, name: string, description: string): Observable<any> {
    return this.http.put<any>(`${this.base}/channels/${channelId}`, { name, description });
  }

  deleteChannel(channelId: number): Observable<any> {
    return this.http.delete<any>(`${this.base}/channels/${channelId}`);
  }

  addChannelMembers(channelId: number, userIds: number[]): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/${channelId}/members`, { user_ids: userIds });
  }

  removeChannelMember(channelId: number, userId: number): Observable<any> {
    return this.http.delete<any>(`${this.base}/channels/${channelId}/members/${userId}`);
  }

  setChannelAdmin(channelId: number, userId: number, isAdmin: boolean): Observable<any> {
    return this.http.put<any>(`${this.base}/channels/${channelId}/members/${userId}/admin`, { is_admin: isAdmin ? 1 : 0 });
  }

  leaveChannel(channelId: number): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/${channelId}/leave`, {});
  }

  sendChannelMessage(channelId: number, message: string, is_file: boolean = false, file_path: string | null = null): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/${channelId}/messages`, { message, is_file, file_path });
  }

  deleteChannelMessage(channelId: number, messageId: number): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/messages/${messageId}/delete`, {});
  }

  reactToChannelMessage(channelId: number, messageId: number, emoji: string): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/messages/${messageId}/react`, { reaction: emoji });
  }



  createChannel(name: string, description: string): Observable<any> {
    return this.http.post<any>(`${this.base}/channels/`, { name, description });
  }

  connectSync(docId: string): Observable<SyncMessage> {
    this.disconnectSync();
    this.currentDocId = docId;
    this.socket = new WebSocket(`${this.wsBase}/ws/${docId}`);

    this.socket.onmessage = (event) => {
      try { this.messageSubject.next(JSON.parse(event.data)); } catch { }
    };
    this.socket.onerror = (err) => console.error('WS error', err);
    this.socket.onclose = () => {
      if (this.currentDocId === docId) setTimeout(() => this.connectSync(docId), 3000);
    };
    return this.messageSubject.asObservable();
  }

  sendUpdate(content: string, title: string, autosave?: boolean): void {
    if (this.socket?.readyState === WebSocket.OPEN) {
      const payload: any = { type: 'update', content, title };
      if (autosave !== undefined) {
        payload.autosave = autosave;
      }
      this.socket.send(JSON.stringify(payload));
    }
  }

  sendCellUpdate(sheetIdx: number, r: number, c: number, value: any, formatting?: any): void {
    if (this.socket?.readyState === WebSocket.OPEN)
      this.socket.send(JSON.stringify({ type: 'cell_update', sheetIdx, r, c, value, formatting }));
  }

  sendCursor(r: number, c: number): void {
    if (this.socket?.readyState === WebSocket.OPEN)
      this.socket.send(JSON.stringify({ type: 'cursor', r, c }));
  }

  disconnectSync(): void {
    this.currentDocId = null;
    if (this.socket) {
      this.socket.onclose = null;
      this.socket.close();
      this.socket = null;
    }
  }

  ngOnDestroy(): void {
    this.disconnectSync();
    this.messageSubject.complete();
  }
}