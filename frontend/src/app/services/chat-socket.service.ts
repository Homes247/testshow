import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
import { environment } from '../../environments/environment';
import { AuthService } from './auth.service';

export interface TypingUser {
  from_user_id: number;
  from_user_name: string;
  from_avatar_color: string;
  from_avatar_url: string | null;
}

@Injectable({
  providedIn: 'root'
})
export class ChatSocketService {
  private ws: WebSocket | null = null;
  public typing$ = new Subject<TypingUser>();
  public stopTyping$ = new Subject<{ from_user_id: number }>();
  public channelTyping$ = new Subject<TypingUser & { channel_id: number }>();
  public channelStopTyping$ = new Subject<{ channel_id: number, from_user_id: number }>();
  public newMessage$ = new Subject<any>();
  public messagesRead$ = new Subject<{ conversation_id: number }>();
  private reconnectTimer: any;
  private pingInterval: any;

  private _typingSent: { [key: string]: boolean } = {};
  private _typingStopTimers: { [key: string]: any } = {};

  constructor(private auth: AuthService) {}

  public connect(token: string) {
    if (this.ws) return;
    if (!token) return;
    const wsUrl = environment.wsUrl + '/ws/chat-hub?token=' + token;
    
    try {
      this.ws = new WebSocket(wsUrl);

      this.ws.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data);
          if (data.type === 'typing') this.typing$.next(data);
          else if (data.type === 'stop_typing') this.stopTyping$.next(data);
          else if (data.type === 'channel_typing') this.channelTyping$.next(data);
          else if (data.type === 'channel_stop_typing') this.channelStopTyping$.next(data);
          else if (data.type === 'messages_read') this.messagesRead$.next(data);
          else if (data.type === 'new_message' || data.type === 'channel_new_message') {
            this.newMessage$.next(data);
          }
        } catch (e) {}
      };

      this.ws.onclose = () => {
        this.ws = null;
        clearInterval(this.pingInterval);
        clearTimeout(this.reconnectTimer);
        this.reconnectTimer = setTimeout(() => this.connect(token), 3000);
      };
      
      this.ws.onerror = () => {
        if (this.ws) {
            this.ws.close();
        }
      };

      this.ws.onopen = () => {
        this.pingInterval = setInterval(() => {
          if (this.ws && this.ws.readyState === WebSocket.OPEN) {
            this.ws.send(JSON.stringify({ type: 'ping' }));
          }
        }, 25000);
      };
    } catch (e) {
      console.error("ChatSocket connection error", e);
    }
  }

  public disconnect() {
    clearTimeout(this.reconnectTimer);
    clearInterval(this.pingInterval);
    if (this.ws) {
      this.ws.onclose = null;
      this.ws.close();
      this.ws = null;
    }
  }

  public emitTyping(toUserId: number) {
    const key = `dm_${toUserId}`;
    if (!this._typingSent[key]) {
      this._typingSent[key] = true;
      this._send({ type: 'typing', to_user_id: toUserId });
    }
    clearTimeout(this._typingStopTimers[key]);
    this._typingStopTimers[key] = setTimeout(() => {
      this._typingSent[key] = false;
      this._send({ type: 'stop_typing', to_user_id: toUserId });
    }, 2000);
  }

  public emitChannelTyping(channelId: number) {
    const key = `ch_${channelId}`;
    if (!this._typingSent[key]) {
      this._typingSent[key] = true;
      this._send({ type: 'channel_typing', channel_id: channelId });
    }
    clearTimeout(this._typingStopTimers[key]);
    this._typingStopTimers[key] = setTimeout(() => {
      this._typingSent[key] = false;
      this._send({ type: 'channel_stop_typing', channel_id: channelId });
    }, 2000);
  }

  public forceStopTyping(toUserId: number) {
    const key = `dm_${toUserId}`;
    clearTimeout(this._typingStopTimers[key]);
    this._typingSent[key] = false;
    this._send({ type: 'stop_typing', to_user_id: toUserId });
  }

  public forceStopChannelTyping(channelId: number) {
    const key = `ch_${channelId}`;
    clearTimeout(this._typingStopTimers[key]);
    this._typingSent[key] = false;
    this._send({ type: 'channel_stop_typing', channel_id: channelId });
  }

  private _send(data: any) {
    if (this.ws && this.ws.readyState === WebSocket.OPEN) {
      this.ws.send(JSON.stringify(data));
    }
  }
}
