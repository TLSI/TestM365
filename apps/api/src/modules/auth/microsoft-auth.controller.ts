import { Controller, Get, Query, Redirect, HttpCode, Delete } from '@nestjs/common';
import { ApiTags, ApiBearerAuth, ApiOperation } from '@nestjs/swagger';
import { MicrosoftAuthService } from './microsoft-auth.service';
import { CurrentUser } from '../../core/auth/decorators/current-user.decorator';
import { Public } from '../../core/auth/decorators/public.decorator';
import { ConfigService } from '@nestjs/config';

@ApiTags('Microsoft Auth')
@Controller('auth/microsoft')
export class MicrosoftAuthController {
  constructor(
    private readonly msAuthService: MicrosoftAuthService,
    private readonly config: ConfigService,
  ) {}

  @ApiBearerAuth()
  @ApiOperation({ summary: 'Get Microsoft OAuth authorization URL' })
  @Get('url')
  getAuthUrl(@CurrentUser() user: { id: string }) {
    return { url: this.msAuthService.getAuthorizationUrl(user.id) };
  }

  @Public()
  @ApiOperation({ summary: 'OAuth callback (receives code from Azure AD)' })
  @Get('callback')
  @Redirect()
  async callback(
    @Query('code') code: string,
    @Query('state') userId: string,
  ) {
    // In production, state would carry a signed token; simplified here
    await this.msAuthService.handleCallback(code, userId);
    const frontendUrl = this.config.get('FRONTEND_URL', 'http://localhost:4200');
    return { url: `${frontendUrl}/microsoft?connected=true` };
  }

  @ApiBearerAuth()
  @ApiOperation({ summary: 'Get current Microsoft 365 connection status' })
  @Get('status')
  getStatus(@CurrentUser() user: { id: string }) {
    return this.msAuthService.getConnectionStatus(user.id);
  }

  @ApiBearerAuth()
  @ApiOperation({ summary: 'Disconnect Microsoft 365 account' })
  @Delete('disconnect')
  @HttpCode(200)
  disconnect(@CurrentUser() user: { id: string }) {
    return this.msAuthService.disconnect(user.id);
  }
}
