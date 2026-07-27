import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { InjectRedis } from '@nestjs-modules/ioredis';
import type { Redis } from 'ioredis';

const APP_TOKEN_CACHE_KEY = 'm365:apptoken';
const APP_TOKEN_TTL_SECONDS = 3480; // 58 minutes (tokens expire in 1h)

@Injectable()
export class MicrosoftGraphService {
  private readonly logger = new Logger(MicrosoftGraphService.name);
  private msalApp: ConfidentialClientApplication | null = null;

  constructor(
    private readonly config: ConfigService,
    @InjectRedis() private readonly redis: Redis,
  ) {
    const clientId = config.get<string>('AZURE_CLIENT_ID');
    if (clientId) {
      this.msalApp = new ConfidentialClientApplication({
        auth: {
          clientId,
          clientSecret: config.get<string>('AZURE_CLIENT_SECRET'),
          authority: `https://login.microsoftonline.com/${config.get('AZURE_TENANT_ID')}`,
        },
      });
    } else {
      this.logger.warn('AZURE_CLIENT_ID not set — Microsoft 365 integration disabled');
    }
  }

  get isConfigured(): boolean {
    return this.msalApp !== null;
  }

  /** Returns a Graph client authenticated as the application (client_credentials).
   *  Used for SharePoint proxy operations. Token is cached in Redis. */
  async getAppClient(): Promise<Client> {
    const token = await this.getAppToken();
    return Client.init({
      authProvider: (done) => done(null, token),
    });
  }

  /** Returns a Graph client authenticated as a specific user (delegated). */
  getUserClient(accessToken: string): Client {
    return Client.init({
      authProvider: (done) => done(null, accessToken),
    });
  }

  private async getAppToken(): Promise<string> {
    const cached = await this.redis.get(APP_TOKEN_CACHE_KEY);
    if (cached) return cached;

    if (!this.msalApp) throw new Error('Microsoft 365 not configured');

    const result = await this.msalApp.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    if (!result?.accessToken) throw new Error('Failed to acquire app token');

    await this.redis.setex(APP_TOKEN_CACHE_KEY, APP_TOKEN_TTL_SECONDS, result.accessToken);
    return result.accessToken;
  }

  /** Exchanges an authorization code for access + refresh tokens (delegated OAuth flow). */
  async exchangeCodeForTokens(code: string, redirectUri: string) {
    if (!this.msalApp) throw new Error('Microsoft 365 not configured');

    return this.msalApp.acquireTokenByCode({
      code,
      redirectUri,
      scopes: [
        'openid',
        'profile',
        'email',
        'offline_access',
        'User.Read',
        'Calendars.ReadWrite',
        'Mail.ReadWrite',
        'Mail.Send',
        'Files.ReadWrite.All',
        'Sites.ReadWrite.All',
      ],
    });
  }

  /** Gets the OAuth2 authorization URL to redirect the user to. `state` round-trips the
   *  requesting userId through Azure AD so the callback can attribute the tokens. */
  getAuthorizationUrl(state: string): string {
    const clientId = this.config.get<string>('AZURE_CLIENT_ID');
    const tenantId = this.config.get<string>('AZURE_TENANT_ID');
    const redirectUri = this.config.get<string>('AZURE_REDIRECT_URI');
    const scopes = [
      'openid',
      'profile',
      'email',
      'offline_access',
      'User.Read',
      'Calendars.ReadWrite',
      'Mail.ReadWrite',
      'Mail.Send',
      'Files.ReadWrite.All',
      'Sites.ReadWrite.All',
    ].join('%20');

    return (
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
      `?client_id=${clientId}` +
      `&response_type=code` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&scope=${scopes}` +
      `&response_mode=query` +
      `&state=${encodeURIComponent(state)}`
    );
  }
}
