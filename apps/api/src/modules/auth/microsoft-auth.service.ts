import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import * as crypto from 'crypto';
import { PrismaService } from '../../core/database/prisma.service';
import { MicrosoftGraphService } from '../../core/microsoft/microsoft-graph.service';
import { AuditLogService } from '../users/audit-log.service';

const ENCRYPTION_KEY_ENV = 'JWT_SECRET'; // reuse as encryption key material

function encrypt(text: string, key: string): string {
  const iv = crypto.randomBytes(16);
  const k = crypto.scryptSync(key, 'salt', 32);
  const cipher = crypto.createCipheriv('aes-256-cbc', k, iv);
  return iv.toString('hex') + ':' + cipher.update(text, 'utf8', 'hex') + cipher.final('hex');
}

function decrypt(text: string, key: string): string {
  const [ivHex, encrypted] = text.split(':');
  const iv = Buffer.from(ivHex, 'hex');
  const k = crypto.scryptSync(key, 'salt', 32);
  const decipher = crypto.createDecipheriv('aes-256-cbc', k, iv);
  return decipher.update(encrypted, 'hex', 'utf8') + decipher.final('utf8');
}

@Injectable()
export class MicrosoftAuthService {
  private readonly logger = new Logger(MicrosoftAuthService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly graphService: MicrosoftGraphService,
    private readonly config: ConfigService,
    private readonly auditLog: AuditLogService,
  ) {}

  getAuthorizationUrl(): string {
    if (!this.graphService.isConfigured) {
      throw new BadRequestException(
        'Microsoft 365 integration not configured. Set AZURE_CLIENT_ID, AZURE_CLIENT_SECRET and AZURE_TENANT_ID.',
      );
    }
    return this.graphService.getAuthorizationUrl();
  }

  async handleCallback(code: string, userId: string) {
    const redirectUri = this.config.get<string>('AZURE_REDIRECT_URI');
    const result = await this.graphService.exchangeCodeForTokens(code, redirectUri);

    if (!result?.accessToken) {
      throw new BadRequestException('Failed to get tokens from Microsoft');
    }

    // Get user profile from Microsoft
    const client = this.graphService.getUserClient(result.accessToken);
    const msUser = await client.api('/me').get();

    const encKey = this.config.get<string>(ENCRYPTION_KEY_ENV);
    const expiresAt = result.expiresOn ?? new Date(Date.now() + 3600 * 1000);

    await this.prisma.microsoftConnection.upsert({
      where: { userId },
      update: {
        microsoftUserId: msUser.id,
        email: msUser.mail ?? msUser.userPrincipalName,
        accessToken: encrypt(result.accessToken, encKey),
        refreshToken: result.account?.idToken
          ? encrypt(result.account.idToken, encKey)
          : encrypt(result.accessToken, encKey),
        tokenExpiresAt: expiresAt,
        scopes: result.scopes ?? [],
        status: 'ACTIVE',
      },
      create: {
        userId,
        microsoftUserId: msUser.id,
        email: msUser.mail ?? msUser.userPrincipalName,
        accessToken: encrypt(result.accessToken, encKey),
        refreshToken: result.account?.idToken
          ? encrypt(result.account.idToken, encKey)
          : encrypt(result.accessToken, encKey),
        tokenExpiresAt: expiresAt,
        scopes: result.scopes ?? [],
      },
    });

    await this.auditLog.log({
      userId,
      action: 'microsoft.connect',
      resource: 'MicrosoftConnection',
      details: { microsoftEmail: msUser.mail },
    });

    this.logger.log(`User ${userId} connected Microsoft account: ${msUser.mail}`);
    return { connected: true, microsoftEmail: msUser.mail ?? msUser.userPrincipalName };
  }

  async disconnect(userId: string) {
    await this.prisma.microsoftConnection.delete({ where: { userId } }).catch(() => null);
    await this.auditLog.log({ userId, action: 'microsoft.disconnect', resource: 'MicrosoftConnection' });
    return { disconnected: true };
  }

  async getConnectionStatus(userId: string) {
    const conn = await this.prisma.microsoftConnection.findUnique({
      where: { userId },
      select: {
        email: true,
        status: true,
        scopes: true,
        syncCalendar: true,
        syncMail: true,
        syncDocuments: true,
        lastSyncAt: true,
        tokenExpiresAt: true,
      },
    });

    return {
      connected: !!conn,
      configured: this.graphService.isConfigured,
      ...conn,
    };
  }

  /** Returns a live Graph client for a user using their stored (decrypted) token. */
  async getClientForUser(userId: string) {
    const conn = await this.prisma.microsoftConnection.findUnique({ where: { userId } });
    if (!conn || conn.status !== 'ACTIVE') {
      throw new BadRequestException('Microsoft account not connected or inactive');
    }
    const encKey = this.config.get<string>(ENCRYPTION_KEY_ENV);
    const token = decrypt(conn.accessToken, encKey);
    return this.graphService.getUserClient(token);
  }
}
