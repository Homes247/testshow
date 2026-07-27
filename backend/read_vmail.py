import re

filepath = r"c:\Users\Homes247\Desktop\office-suite\frontend\src\app\components\chat-widget\chat-widget.component.ts"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace the CHANNELS VIEW part of the template
channels_view = """      <!-- CHANNELS VIEW -->
      <ng-container *ngIf="activeWidget === 'channels' && !openedChannelId">
        <div class="wp-header-white">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined" style="color:#6366f1;">hub</span>
            <span style="font-size:15px; font-weight:600; color:#202124;">Channels</span>
          </div>
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined icon-btn" style="color:#6366f1;">add</span>
            <span class="material-symbols-outlined icon-btn" (click)="close.emit()">close</span>
          </div>
        </div>
        <div class="wp-search">
          <span class="material-symbols-outlined">search</span>
          <input type="text" placeholder="Search channels..." [(ngModel)]="searchQuery">
        </div>
        <div class="wp-section-title">MY CHANNELS</div>
        <div class="wp-body-list">
          <div class="chat-list-item" *ngFor="let c of filteredChannels()" (click)="openChannel(c.id, c.name)">
            <div class="cli-av" style="background:#6366f1">#</div>
            <div class="cli-info">
              <div class="cli-top">
                <span class="cli-name">{{ c.name }}</span>
                <span class="cli-time">{{ formatTime(c.last_time) }}</span>
              </div>
              <div class="cli-preview">{{ c.last_message || 'No messages' }}</div>
            </div>
          </div>
          <div *ngIf="channels.length === 0" style="padding: 20px; text-align:center; color:#9aa0a6; font-size:13px;">
            No channels
          </div>
        </div>
      </ng-container>

      <!-- OPENED CHANNEL VIEW -->
      <ng-container *ngIf="activeWidget === 'channels' && openedChannelId">
        <div class="wp-header-white" style="background:#f8f9fa;">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined icon-btn" (click)="closeChannel()">arrow_back</span>
            <div class="cli-av" style="width:32px; height:32px; font-size:16px; margin-right:4px; background:#6366f1;">#</div>
            <div style="display:flex; flex-direction:column;">
                <span style="font-size:14px; font-weight:600; color:#202124;">{{ openedChannelName }}</span>
                <span style="font-size:11px; color:#5f6368;">Channel</span>
            </div>
          </div>
        </div>
        <div class="wp-body-list" style="padding:12px; display:flex; flex-direction:column; gap:8px; flex:1; background:#f1f3f4;" #scrollMe>
            <div *ngFor="let msg of messages" style="display:flex; flex-direction:column; max-width:80%;" [style.align-self]="msg.is_mine ? 'flex-end' : 'flex-start'">
                <span *ngIf="!msg.is_mine" style="font-size:11px; color:#5f6368; margin-bottom:2px; margin-left:4px;">{{ msg.sender_name }}</span>
                <div style="padding:8px 12px; border-radius:12px; font-size:13px; line-height:1.4; word-break:break-word;" [style.background]="msg.is_mine ? '#6a35ff' : '#fff'" [style.color]="msg.is_mine ? '#fff' : '#202124'">
                    {{ msg.message }}
                </div>
                <div style="font-size:10px; color:#9aa0a6; margin-top:2px; display:flex; justify-content:flex-end;">{{ formatTime(msg.created_at) }}</div>
            </div>
        </div>
        <div style="padding:8px; border-top:1px solid #e0e0e0; display:flex; gap:8px; background:#fff;">
           <input type="text" placeholder="Type a message..." [(ngModel)]="newMessage" (keydown.enter)="sendMessage()" style="flex:1; border:1px solid #dadce0; border-radius:20px; padding:8px 14px; outline:none; font-size:13px;">
           <span class="material-symbols-outlined" style="color:#6a35ff; cursor:pointer; align-self:center; font-size:24px;" (click)="sendMessage()">send</span>
        </div>
      </ng-container>"""

content = re.sub(r"<!-- CHANNELS VIEW -->.*?<!-- CONTACTS VIEW -->", channels_view + "\n\n      <!-- CONTACTS VIEW -->", content, flags=re.DOTALL)

# Add logic properties
if "openedChannelId: number | null = null;" not in content:
    content = content.replace("openedChatUser: any = null;", "openedChatUser: any = null;\n  openedChannelId: number | null = null;\n  openedChannelName: string = '';")

# Update pollData
poll_data_replacement = """  pollData() {
    if (this.activeWidget === 'chat') {
        if (this.openedChatId) {
            this.api.getChatMessages(this.openedChatId).subscribe(res => {
                if (res && res.messages) {
                    this.messages = res.messages;
                    this.openedChatUser = res.other_user;
                }
            });
        } else {
            this.api.getChatConversations().subscribe(res => this.conversations = res);
        }
    } else if (this.activeWidget === 'channels') {
        if (this.openedChannelId) {
            this.api.getChannelMessages(this.openedChannelId).subscribe(res => {
                if (res && res.messages) {
                    this.messages = res.messages;
                }
            });
        } else {
            this.api.getChatChannels().subscribe(res => this.channels = res);
        }
    } else if (this.activeWidget === 'contacts') {
        this.api.getChatContacts().subscribe(res => this.contacts = res);
    }
  }"""
content = re.sub(r"pollData\(\) \{.*?\n  \}", poll_data_replacement, content, flags=re.DOTALL)

# Update sendMessage
send_message_replacement = """  sendMessage() {
      if (!this.newMessage.trim()) return;
      if (this.activeWidget === 'chat' && this.openedChatId) {
          this.api.sendChatMessage(this.openedChatId, this.newMessage).subscribe(() => {
              this.newMessage = '';
              this.pollData();
          });
      } else if (this.activeWidget === 'channels' && this.openedChannelId) {
          this.api.sendChannelMessage(this.openedChannelId, this.newMessage).subscribe(() => {
              this.newMessage = '';
              this.pollData();
          });
      }
  }"""
content = re.sub(r"sendMessage\(\) \{.*?\n  \}", send_message_replacement, content, flags=re.DOTALL)

# Add openChannel and closeChannel
if "openChannel(" not in content:
    open_channel_code = """
  openChannel(channelId: number, channelName: string) {
      this.openedChannelId = channelId;
      this.openedChannelName = channelName;
      this.messages = [];
      this.searchQuery = '';
      this.pollData();
  }

  closeChannel() {
      this.openedChannelId = null;
      this.openedChannelName = '';
      this.pollData();
  }
"""
    content = content.replace("closeChat() {", open_channel_code + "\n  closeChat() {")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print("done")
