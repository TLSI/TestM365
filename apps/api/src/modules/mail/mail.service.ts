import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../../core/database/prisma.service';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AuditLogService } from '../users/audit-log.service';
import { SendEmailDto } from './dto/send-email.dto';

@Injectable()
export class MailService {
  private readonly logger = new Logger(MailService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly msAuth: MicrosoftAuthService,
    private readonly auditLog: AuditLogService,
  ) {}

  /** Send an email via Graph API /me/sendMail using the user's Outlook account */
  async sendEmail(userId: string, dto: SendEmailDto) {
    const client = await this.msAuth.getClientForUser(userId);

    const message = {
      subject: dto.subject,
      body: { contentType: 'HTML', content: dto.body },
      toRecipients: dto.to.map((addr) => ({ emailAddress: { address: addr } })),
      ccRecipients: (dto.cc ?? []).map((addr) => ({ emailAddress: { address: addr } })),
    };

    await client.api('/me/sendMail').post({ message, saveToSentItems: true });

    const user = await this.prisma.user.findUnique({ where: { id: userId }, select: { email: true } });

    const record = await this.prisma.emailRecord.create({
      data: {
        userId,
        clientId: dto.clientId ?? null,
        direction: 'OUTBOUND',
        from: user?.email ?? '',
        to: dto.to,
        cc: dto.cc ?? [],
        subject: dto.subject,
        bodyPreview: dto.body.substring(0, 200),
        sentAt: new Date(),
      },
    });

    await this.auditLog.log({ userId, action: 'mail.send', resource: 'EmailRecord', resourceId: record.id });
    return record;
  }

  /** List recent inbox messages from Outlook */
  async listInbox(userId: string, top = 20) {
    const client = await this.msAuth.getClientForUser(userId);

    const response = await client
      .api('/me/mailFolders/inbox/messages')
      .query({
        $select: 'id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments',
        $top: top,
        $orderby: 'receivedDateTime desc',
      })
      .get();

    return response?.value ?? [];
  }

  /** Link an existing Outlook email (by message ID) to a local client record */
  async linkEmail(userId: string, messageId: string, clientId: string) {
    const client = await this.msAuth.getClientForUser(userId);

    const msg = await client
      .api(`/me/messages/${messageId}`)
      .select('id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments')
      .get();

    const record = await this.prisma.emailRecord.create({
      data: {
        externalId: msg.id,
        userId,
        clientId,
        direction: 'INBOUND',
        from: msg.from?.emailAddress?.address ?? '',
        to: (msg.toRecipients ?? []).map((r: { emailAddress: { address: string } }) => r.emailAddress.address),
        subject: msg.subject ?? '',
        bodyPreview: msg.bodyPreview ?? '',
        hasAttachments: msg.hasAttachments ?? false,
        sentAt: new Date(msg.receivedDateTime),
      },
    });

    await this.auditLog.log({ userId, action: 'mail.link', resource: 'EmailRecord', resourceId: record.id, details: { clientId } });
    return record;
  }

  /** Get email records linked to a specific client */
  async getLinkedEmails(clientId: string) {
    return this.prisma.emailRecord.findMany({
      where: { clientId },
      orderBy: { sentAt: 'desc' },
    });
  }

  /** Get all local email records for a user */
  async getUserEmails(userId: string) {
    return this.prisma.emailRecord.findMany({
      where: { userId },
      orderBy: { sentAt: 'desc' },
      take: 50,
    });
  }
}
