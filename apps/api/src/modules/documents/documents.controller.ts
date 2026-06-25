import {
  Controller, Get, Post, Param, Query, UploadedFile,
  UseInterceptors, Res, HttpCode,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { ApiTags, ApiBearerAuth, ApiOperation, ApiConsumes } from '@nestjs/swagger';
import { Response } from 'express';
import { DocumentsService } from './documents.service';
import { CurrentUser } from '../../core/auth/decorators/current-user.decorator';

interface UploadedFile {
  originalname: string;
  buffer: Buffer;
  mimetype: string;
  size: number;
}

@ApiTags('Documents')
@ApiBearerAuth()
@Controller('documents')
export class DocumentsController {
  constructor(private readonly documentsService: DocumentsService) {}

  @ApiOperation({ summary: 'Create SharePoint site + folder structure for a client' })
  @Post('clients/:clientId/site')
  @HttpCode(200)
  createSite(@Param('clientId') clientId: string, @CurrentUser() user: { id: string }) {
    return this.documentsService.createClientSite(clientId, user.id);
  }

  @ApiOperation({ summary: 'List documents for a client from SharePoint' })
  @Get('clients/:clientId')
  list(@Param('clientId') clientId: string, @Query('folder') folder = '') {
    return this.documentsService.listDocuments(clientId, folder);
  }

  @ApiOperation({ summary: 'Upload file to SharePoint for a client' })
  @ApiConsumes('multipart/form-data')
  @Post('clients/:clientId/upload')
  @UseInterceptors(FileInterceptor('file'))
  async upload(
    @Param('clientId') clientId: string,
    @CurrentUser() user: { id: string },
    @UploadedFile() file: UploadedFile,
    @Query('folder') folder = '',
  ) {
    return this.documentsService.uploadDocument(clientId, user.id, file.originalname, file.buffer, folder);
  }

  @ApiOperation({ summary: 'Get embeddable preview URL (iframe, no M365 account needed)' })
  @Get('clients/:clientId/items/:itemId/preview')
  preview(@Param('clientId') clientId: string, @Param('itemId') itemId: string) {
    return this.documentsService.getPreviewUrl(clientId, itemId);
  }

  @ApiOperation({ summary: 'Get temporary edit link (Office Online, anonymous, expires in 4h)' })
  @Get('clients/:clientId/items/:itemId/edit-link')
  editLink(@Param('clientId') clientId: string, @Param('itemId') itemId: string) {
    return this.documentsService.getEditLink(clientId, itemId);
  }

  @ApiOperation({ summary: 'Download file via proxy (client never sees SharePoint URL)' })
  @Get('clients/:clientId/items/:itemId/download')
  async download(
    @Param('clientId') clientId: string,
    @Param('itemId') itemId: string,
    @Res() res: Response,
  ) {
    const buffer = await this.documentsService.downloadFile(clientId, itemId);
    res.set({ 'Content-Type': 'application/octet-stream' });
    res.send(buffer);
  }
}
