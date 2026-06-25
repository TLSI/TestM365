import { Module } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { APP_GUARD } from '@nestjs/core';
import { ThrottlerModule, ThrottlerGuard } from '@nestjs/throttler';
import { RedisModule } from '@nestjs-modules/ioredis';
import { PrismaModule } from '../core/database/prisma.module';
import { MicrosoftModule } from '../core/microsoft/microsoft.module';
import { JwtAuthGuard } from '../core/auth/guards/jwt-auth.guard';
import { AuthModule } from '../modules/auth/auth.module';
import { UsersModule } from '../modules/users/users.module';
import { ClientsModule } from '../modules/clients/clients.module';
import { CalendarModule } from '../modules/calendar/calendar.module';
import { MailModule } from '../modules/mail/mail.module';
import { DocumentsModule } from '../modules/documents/documents.module';
import { TeamsModule } from '../modules/teams/teams.module';

@Module({
  imports: [
    ConfigModule.forRoot({ isGlobal: true }),
    ThrottlerModule.forRoot([{ ttl: 60000, limit: 100 }]),
    RedisModule.forRootAsync({
      inject: [ConfigService],
      useFactory: (config: ConfigService) => ({
        type: 'single',
        url: `redis://${config.get('REDIS_HOST', 'localhost')}:${config.get('REDIS_PORT', 6380)}`,
      }),
    }),
    PrismaModule,
    MicrosoftModule,
    AuthModule,
    UsersModule,
    ClientsModule,
    CalendarModule,
    MailModule,
    DocumentsModule,
    TeamsModule,
  ],
  providers: [
    { provide: APP_GUARD, useClass: JwtAuthGuard },
    { provide: APP_GUARD, useClass: ThrottlerGuard },
  ],
})
export class AppModule {}
