import { Body, Controller, Delete, Get, Param, Post, HttpCode } from '@nestjs/common';
import { ApiTags, ApiBearerAuth, ApiOperation } from '@nestjs/swagger';
import { CalendarService } from './calendar.service';
import { CreateEventDto } from './dto/create-event.dto';
import { CurrentUser } from '../../core/auth/decorators/current-user.decorator';

@ApiTags('Calendar')
@ApiBearerAuth()
@Controller('calendar')
export class CalendarController {
  constructor(private readonly calendarService: CalendarService) {}

  @ApiOperation({ summary: 'Get all local calendar events' })
  @Get('events')
  getEvents(@CurrentUser() user: { id: string }) {
    return this.calendarService.getEvents(user.id);
  }

  @ApiOperation({ summary: 'Sync events from Outlook (initial + delta)' })
  @Post('sync')
  @HttpCode(200)
  sync(@CurrentUser() user: { id: string }) {
    return this.calendarService.syncFromOutlook(user.id);
  }

  @ApiOperation({ summary: 'Create event (pushes to Outlook + Teams link if online)' })
  @Post('events')
  createEvent(@CurrentUser() user: { id: string }, @Body() dto: CreateEventDto) {
    return this.calendarService.createEvent(user.id, dto);
  }

  @ApiOperation({ summary: 'Delete event (removes from Outlook too)' })
  @Delete('events/:id')
  deleteEvent(@CurrentUser() user: { id: string }, @Param('id') id: string) {
    return this.calendarService.deleteEvent(user.id, id);
  }
}
