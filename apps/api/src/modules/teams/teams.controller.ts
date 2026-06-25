import { Body, Controller, Post } from '@nestjs/common';
import { ApiTags, ApiBearerAuth, ApiOperation } from '@nestjs/swagger';
import { TeamsService } from './teams.service';
import { SendTeamsNotificationDto } from './dto/send-teams-notification.dto';
import { CurrentUser } from '../../core/auth/decorators/current-user.decorator';
import { IsISO8601, IsString } from 'class-validator';
import { ApiProperty } from '@nestjs/swagger';

class CreateMeetingDto {
  @ApiProperty()
  @IsString()
  subject: string;

  @ApiProperty()
  @IsISO8601()
  startTime: string;

  @ApiProperty()
  @IsISO8601()
  endTime: string;
}

@ApiTags('Teams')
@ApiBearerAuth()
@Controller('teams')
export class TeamsController {
  constructor(private readonly teamsService: TeamsService) {}

  @ApiOperation({ summary: 'Send notification to Teams channel via incoming webhook' })
  @Post('notify')
  notify(@CurrentUser() user: { id: string }, @Body() dto: SendTeamsNotificationDto) {
    return this.teamsService.sendChannelNotification(user.id, dto);
  }

  @ApiOperation({ summary: 'Create a Teams online meeting' })
  @Post('meeting')
  meeting(@CurrentUser() user: { id: string }, @Body() dto: CreateMeetingDto) {
    return this.teamsService.createMeeting(user.id, dto.subject, dto.startTime, dto.endTime);
  }
}
