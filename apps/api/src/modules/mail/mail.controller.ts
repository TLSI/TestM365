import { Body, Controller, Get, Param, Post, Query } from '@nestjs/common';
import { ApiTags, ApiBearerAuth, ApiOperation } from '@nestjs/swagger';
import { MailService } from './mail.service';
import { SendEmailDto } from './dto/send-email.dto';
import { CurrentUser } from '../../core/auth/decorators/current-user.decorator';

@ApiTags('Mail')
@ApiBearerAuth()
@Controller('mail')
export class MailController {
  constructor(private readonly mailService: MailService) {}

  @ApiOperation({ summary: 'Send email via Outlook (Graph /me/sendMail)' })
  @Post('send')
  send(@CurrentUser() user: { id: string }, @Body() dto: SendEmailDto) {
    return this.mailService.sendEmail(user.id, dto);
  }

  @ApiOperation({ summary: 'List recent inbox messages from Outlook' })
  @Get('inbox')
  inbox(@CurrentUser() user: { id: string }, @Query('top') top = 20) {
    return this.mailService.listInbox(user.id, +top);
  }

  @ApiOperation({ summary: 'Link an Outlook message to a client' })
  @Post('link/:messageId')
  link(
    @CurrentUser() user: { id: string },
    @Param('messageId') messageId: string,
    @Query('clientId') clientId: string,
  ) {
    return this.mailService.linkEmail(user.id, messageId, clientId);
  }

  @ApiOperation({ summary: 'Get my local email records' })
  @Get('records')
  records(@CurrentUser() user: { id: string }) {
    return this.mailService.getUserEmails(user.id);
  }

  @ApiOperation({ summary: 'Get emails linked to a client' })
  @Get('records/client/:clientId')
  clientEmails(@Param('clientId') clientId: string) {
    return this.mailService.getLinkedEmails(clientId);
  }
}
