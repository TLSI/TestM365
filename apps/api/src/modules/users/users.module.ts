import { Module } from '@nestjs/common';
import { AuditLogService } from './audit-log.service';
import { UsersController } from './users.controller';
import { UsersService } from './users.service';

@Module({
  controllers: [UsersController],
  providers: [UsersService, AuditLogService],
  exports: [AuditLogService, UsersService],
})
export class UsersModule {}
