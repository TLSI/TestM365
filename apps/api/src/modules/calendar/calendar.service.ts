import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../../core/database/prisma.service';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AuditLogService } from '../users/audit-log.service';
import { CreateEventDto } from './dto/create-event.dto';

@Injectable()
export class CalendarService {
  private readonly logger = new Logger(CalendarService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly msAuth: MicrosoftAuthService,
    private readonly auditLog: AuditLogService,
  ) {}

  /** Pull all events from Outlook for the next 90 days and store locally */
  async syncFromOutlook(userId: string): Promise<{ synced: number }> {
    const client = await this.msAuth.getClientForUser(userId);

    const now = new Date();
    const end = new Date();
    end.setDate(end.getDate() + 90);

    const response = await client
      .api('/me/calendarView')
      .query({
        startDateTime: now.toISOString(),
        endDateTime: end.toISOString(),
        $select: 'id,subject,body,start,end,location,isOnlineMeeting,onlineMeeting,attendees,isAllDay,categories',
        $top: 50,
      })
      .get();

    const events: unknown[] = response?.value ?? [];
    let synced = 0;

    for (const evt of events as Record<string, unknown>[]) {
      await this.prisma.calendarEvent.upsert({
        where: { externalId: evt['id'] as string },
        update: {
          title: evt['subject'] as string,
          startTime: new Date((evt['start'] as { dateTime: string })?.dateTime),
          endTime: new Date((evt['end'] as { dateTime: string })?.dateTime),
          location: (evt['location'] as { displayName?: string })?.displayName ?? null,
          isOnline: (evt['isOnlineMeeting'] as boolean) ?? false,
          meetingUrl: (evt['onlineMeeting'] as { joinUrl?: string })?.joinUrl ?? null,
          attendees: evt['attendees'] ?? [],
          allDay: (evt['isAllDay'] as boolean) ?? false,
          source: 'OUTLOOK',
          synced: true,
        },
        create: {
          externalId: evt['id'] as string,
          userId,
          title: (evt['subject'] as string) || '(Sin título)',
          startTime: new Date((evt['start'] as { dateTime: string })?.dateTime),
          endTime: new Date((evt['end'] as { dateTime: string })?.dateTime),
          location: (evt['location'] as { displayName?: string })?.displayName ?? null,
          isOnline: (evt['isOnlineMeeting'] as boolean) ?? false,
          meetingUrl: (evt['onlineMeeting'] as { joinUrl?: string })?.joinUrl ?? null,
          attendees: evt['attendees'] ?? [],
          allDay: (evt['isAllDay'] as boolean) ?? false,
          source: 'OUTLOOK',
          synced: true,
        },
      });
      synced++;
    }

    await this.prisma.microsoftConnection.update({
      where: { userId },
      data: { lastSyncAt: new Date() },
    });

    return { synced };
  }

  /** Create an event in FORLOPD and push it to Outlook */
  async createEvent(userId: string, dto: CreateEventDto) {
    const client = await this.msAuth.getClientForUser(userId);

    const outlookEvent = {
      subject: dto.title,
      body: { contentType: 'Text', content: dto.description ?? '' },
      start: { dateTime: dto.startTime, timeZone: 'UTC' },
      end: { dateTime: dto.endTime, timeZone: 'UTC' },
      location: dto.location ? { displayName: dto.location } : undefined,
      isOnlineMeeting: dto.isOnline ?? false,
      onlineMeetingProvider: dto.isOnline ? 'teamsForBusiness' : undefined,
      attendees: (dto.attendees ?? []).map((email: string) => ({
        emailAddress: { address: email },
        type: 'required',
      })),
    };

    const created = await client.api('/me/events').post(outlookEvent);

    const local = await this.prisma.calendarEvent.create({
      data: {
        externalId: created.id,
        userId,
        title: dto.title,
        description: dto.description,
        startTime: new Date(dto.startTime),
        endTime: new Date(dto.endTime),
        location: dto.location,
        isOnline: dto.isOnline ?? false,
        meetingUrl: created.onlineMeeting?.joinUrl ?? null,
        attendees: dto.attendees ?? [],
        source: 'PLATFORM',
        synced: true,
      },
    });

    await this.auditLog.log({ userId, action: 'calendar.create', resource: 'CalendarEvent', resourceId: local.id });
    return local;
  }

  /** Delete an event locally and in Outlook */
  async deleteEvent(userId: string, eventId: string) {
    const event = await this.prisma.calendarEvent.findFirst({ where: { id: eventId, userId } });
    if (!event) return { deleted: false };

    if (event.externalId) {
      const client = await this.msAuth.getClientForUser(userId);
      await client.api(`/me/events/${event.externalId}`).delete().catch((e) =>
        this.logger.warn(`Could not delete Outlook event: ${e.message}`),
      );
    }

    await this.prisma.calendarEvent.delete({ where: { id: eventId } });
    await this.auditLog.log({ userId, action: 'calendar.delete', resource: 'CalendarEvent', resourceId: eventId });
    return { deleted: true };
  }

  /** Get all local events for a user */
  async getEvents(userId: string) {
    return this.prisma.calendarEvent.findMany({
      where: { userId },
      orderBy: { startTime: 'asc' },
    });
  }
}
