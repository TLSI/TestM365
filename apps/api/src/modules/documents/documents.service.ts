import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import { PrismaService } from '../../core/database/prisma.service';
import { MicrosoftGraphService } from '../../core/microsoft/microsoft-graph.service';
import { AuditLogService } from '../users/audit-log.service';

@Injectable()
export class DocumentsService {
  private readonly logger = new Logger(DocumentsService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly graphService: MicrosoftGraphService,
    private readonly config: ConfigService,
    private readonly auditLog: AuditLogService,
  ) {}

  // ─── SharePoint Site Management ──────────────────────────────────────────────

  /** Creates a SharePoint site for a client and stores the IDs in the DB */
  async createClientSite(clientId: string, userId: string) {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    const owner = await this.prisma.user.findUniqueOrThrow({ where: { id: userId }, select: { email: true } });
    const appClient = await this.graphService.getAppClient();

    const tenantDomain = this.config.get<string>('SHAREPOINT_TENANT_DOMAIN');
    const siteName = `cliente-${client.name.toLowerCase().replace(/[^a-z0-9]/g, '-')}`;

    const response = await appClient.api(`/sites/${tenantDomain}:/sites/${siteName}?$select=id,displayName,webUrl`).get().catch(() => null);

    const siteId: string =
      response?.id ??
      (await this.provisionSharePointSite(appClient, tenantDomain, siteName, client.name, owner.email));

    // Get the default drive
    const drive = await appClient.api(`/sites/${siteId}/drive`).get();
    const driveId = drive.id;

    // Create standard folder structure
    const folders = ['LOPD', 'Contratos', 'Correspondencia', 'Informes'];
    for (const folder of folders) {
      await appClient
        .api(`/sites/${siteId}/drive/root/children`)
        .post({ name: folder, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' })
        .catch(() => null);
    }

    await this.prisma.client.update({
      where: { id: clientId },
      data: { sharepointSiteId: siteId, sharepointDriveId: driveId },
    });

    await this.auditLog.log({ userId, action: 'document.site_created', resource: 'Client', resourceId: clientId, details: { siteId } });

    return { siteId, driveId, siteName };
  }

  /** Provisions a new SharePoint site collection.
   *  Graph's v1.0 `/sites/root/sites` only creates classic sub-webs, not real site collections —
   *  creating a site collection requires the beta `POST /sites` API (permission `Sites.Create.All`),
   *  which is asynchronous: it returns 202 with an operation to poll until it succeeds. */
  private async provisionSharePointSite(
    appClient: Client,
    tenantDomain: string,
    siteName: string,
    clientName: string,
    ownerEmail: string,
  ): Promise<string> {
    const requestBody = {
      name: siteName,
      webUrl: `https://${tenantDomain}/sites/${siteName}`,
      locale: 'es-ES',
      shareByEmailEnabled: false,
      description: `Documentos del cliente ${clientName}`,
      template: 'sts',
      // Required when calling with application (app-only) permissions: there's no
      // signed-in user Graph can default the owner to.
      ownerIdentityToResolve: { email: ownerEmail },
    };

    // SharePoint site provisioning occasionally fails transiently (observed 500
    // "Error happened while site provisioning") — retry a couple of times before giving up.
    let raw: Response | undefined;
    let lastError = '';
    for (let attempt = 0; attempt < 3; attempt++) {
      raw = (await appClient.api('/sites').version('beta').responseType(ResponseType.RAW).post(requestBody)) as Response;
      if (raw.status === 202) break;
      lastError = `HTTP ${raw.status}: ${await raw.text().catch(() => '<sin cuerpo>')}`;
      if (raw.status < 500) break; // don't retry non-transient errors (bad request, permissions, etc.)
      await new Promise((resolve) => setTimeout(resolve, 3000));
    }

    if (!raw || raw.status !== 202) {
      throw new Error(`Fallo creando site de SharePoint tras reintentos: ${lastError}`);
    }

    const operationId = raw.headers.get('location')?.match(/operationId='([^']+)'/)?.[1];
    if (!operationId) {
      throw new Error('SharePoint devolvió 202 pero sin cabecera Location con el ID de operación');
    }

    for (let attempt = 0; attempt < 30; attempt++) {
      const op = await appClient
        .api(`/sites/getOperationStatus(operationId='${operationId}')`)
        .version('beta')
        .get();

      if (op.status === 'succeeded') {
        // resourceId is a bare site GUID here; normalize to the full `hostname,siteId,webId`
        // form used everywhere else so lookups stay consistent.
        const site = await appClient.api(`/sites/${op.resourceId}?$select=id`).get();
        return site.id;
      }
      if (op.status === 'failed') {
        throw new Error(`La creación del site de SharePoint falló: ${JSON.stringify(op)}`);
      }

      await new Promise((resolve) => setTimeout(resolve, 2000));
    }

    throw new Error('Tiempo de espera agotado creando el site de SharePoint');
  }

  // ─── Document Operations ─────────────────────────────────────────────────────

  /** List documents for a client from SharePoint */
  async listDocuments(clientId: string, folderPath = '') {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    if (!client.sharepointSiteId || !client.sharepointDriveId) {
      return this.prisma.document.findMany({ where: { clientId } });
    }

    const appClient = await this.graphService.getAppClient();
    const endpoint = folderPath
      ? `/sites/${client.sharepointSiteId}/drive/root:/${folderPath}:/children`
      : `/sites/${client.sharepointSiteId}/drive/root/children`;

    const response = await appClient.api(endpoint)
      .select('id,name,size,file,folder,lastModifiedDateTime,webUrl')
      .get();

    return response?.value ?? [];
  }

  /** Upload a file to SharePoint for a client */
  async uploadDocument(
    clientId: string,
    userId: string,
    fileName: string,
    content: Buffer,
    folderPath = '',
  ) {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    if (!client.sharepointSiteId) throw new NotFoundException('No SharePoint site for this client');

    const appClient = await this.graphService.getAppClient();
    const path = folderPath ? `${folderPath}/${fileName}` : fileName;

    const uploaded = await appClient
      .api(`/sites/${client.sharepointSiteId}/drive/root:/${path}:/content`)
      .put(content);

    const doc = await this.prisma.document.create({
      data: {
        name: fileName,
        clientId,
        spItemId: uploaded.id,
        spWebUrl: uploaded.webUrl,
        spDriveId: client.sharepointDriveId ?? undefined,
        folderPath,
        mimeType: uploaded.file?.mimeType,
        size: uploaded.size,
        uploadedByUserId: userId,
      },
    });

    await this.auditLog.log({ userId, action: 'document.upload', resource: 'Document', resourceId: doc.id });
    return doc;
  }

  /** Get a document preview URL from Graph API (embeddable iframe, no M365 account needed) */
  async getPreviewUrl(clientId: string, itemId: string) {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    if (!client.sharepointSiteId) throw new NotFoundException('No SharePoint site for this client');

    const appClient = await this.graphService.getAppClient();
    const preview = await appClient
      .api(`/sites/${client.sharepointSiteId}/drive/items/${itemId}/preview`)
      .post({});

    return { previewUrl: preview.getUrl };
  }

  /** Generate a temporary anonymous edit link (Office Online) — no guest users */
  async getEditLink(clientId: string, itemId: string, expiresInHours = 4) {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    if (!client.sharepointSiteId) throw new NotFoundException('No SharePoint site for this client');

    const appClient = await this.graphService.getAppClient();
    const expiry = new Date(Date.now() + expiresInHours * 3600 * 1000).toISOString();

    const link = await appClient
      .api(`/sites/${client.sharepointSiteId}/drive/items/${itemId}/createLink`)
      .post({ type: 'edit', scope: 'anonymous', expirationDateTime: expiry });

    return { editUrl: link.link?.webUrl, expiresAt: expiry };
  }

  /** Proxy download — streams file content through API, client never sees SharePoint URL */
  async downloadFile(clientId: string, itemId: string): Promise<Buffer> {
    const client = await this.prisma.client.findUniqueOrThrow({ where: { id: clientId } });
    if (!client.sharepointSiteId) throw new NotFoundException('No SharePoint site for this client');

    const appClient = await this.graphService.getAppClient();
    const content = await appClient
      .api(`/sites/${client.sharepointSiteId}/drive/items/${itemId}/content`)
      .getStream();

    const chunks: Uint8Array[] = [];
    for await (const chunk of content) {
      chunks.push(chunk);
    }
    return Buffer.concat(chunks);
  }
}
