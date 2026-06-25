import { Injectable, Logger } from '@nestjs/common';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AuditLogService } from '../users/audit-log.service';
import { SendTeamsNotificationDto } from './dto/send-teams-notification.dto';
import axios from 'axios';

@Injectable()
export class TeamsService {
  private readonly logger = new Logger(TeamsService.name);

  constructor(
    private readonly msAuth: MicrosoftAuthService,
    private readonly auditLog: AuditLogService,
  ) {}

  /** Send a notification to a Teams channel via an incoming webhook URL */
  async sendChannelNotification(userId: string, dto: SendTeamsNotificationDto) {
    const payload = {
      '@type': 'MessageCard',
      '@context': 'http://schema.org/extensions',
      themeColor: '0076D7',
      summary: dto.title,
      sections: [
        {
          activityTitle: dto.title,
          activitySubtitle: dto.subtitle ?? 'TestM365 Platform',
          activityText: dto.body,
          facts: (dto.facts ?? []).map((f) => ({ name: f.name, value: f.value })),
        },
      ],
    };

    await axios.post(dto.webhookUrl, payload);

    await this.auditLog.log({
      userId,
      action: 'teams.notification_sent',
      resource: 'TeamsChannel',
      details: { title: dto.title, webhookUrl: dto.webhookUrl },
    });

    return { sent: true };
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
