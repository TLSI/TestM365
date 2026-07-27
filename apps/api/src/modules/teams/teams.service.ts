import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AuditLogService } from '../users/audit-log.service';
import { SendTeamsNotificationDto } from './dto/send-teams-notification.dto';

@Injectable()
export class TeamsService {
  private readonly logger = new Logger(TeamsService.name);

  constructor(
    private readonly msAuth: MicrosoftAuthService,
    private readonly auditLog: AuditLogService,
  ) {}

  /** Send a notification to a Teams channel via Graph API.
   *  Uses the user's delegated token (ChannelMessage.Send delegated scope). */
  async sendChannelNotification(userId: string, dto: SendTeamsNotificationDto) {
    const { teamId, channelId } = this.parseChannelUrl(dto.channelUrl);
    const appClient = await this.msAuth.getClientForUser(userId);

    let html = `<h2>${dto.title}</h2>`;
    if (dto.subtitle) html += `<p><em>${dto.subtitle}</em></p>`;
    html += `<p>${dto.body}</p>`;
    if (dto.facts?.length) {
      html += '<ul>' + dto.facts.map((f) => `<li><strong>${f.name}:</strong> ${f.value}</li>`).join('') + '</ul>';
    }

    await appClient.api(`/teams/${teamId}/channels/${channelId}/messages`).post({
      body: { contentType: 'html', content: html },
    });

    await this.auditLog.log({
      userId,
      action: 'teams.notification_sent',
      resource: 'TeamsChannel',
      details: { title: dto.title, teamId, channelId },
    });

    return { sent: true };
  }

  private parseChannelUrl(url: string): { teamId: string; channelId: string } {
    try {
      const parsed = new URL(url);
      const teamId = parsed.searchParams.get('groupId');
      const parts = parsed.pathname.split('/');
      const idx = parts.indexOf('channel');
      const channelId = idx !== -1 ? decodeURIComponent(parts[idx + 1]) : null;
      if (!teamId || !channelId) throw new Error();
      return { teamId, channelId };
    } catch {
      throw new BadRequestException(
        'URL de canal no válida. Copia el enlace desde Teams → canal → "Obtener vínculo al canal".',
      );
    }
  }

  /** Create a Teams online meeting via Graph API */
  async createMeeting(userId: string, subject: string, startTime: string, endTime: string) {
    const client = await this.msAuth.getClientForUser(userId);

    const meeting = await client.api('/me/onlineMeetings').post({
      subject,
      startDateTime: startTime,
      endDateTime: endTime,
    });

    await this.auditLog.log({ userId, action: 'teams.meeting_created', resource: 'TeamsMeeting', details: { subject } });

    return {
      id: meeting.id,
      joinUrl: meeting.joinWebUrl,
      subject: meeting.subject,
      startDateTime: meeting.startDateTime,
      endDateTime: meeting.endDateTime,
    };
  }
}
