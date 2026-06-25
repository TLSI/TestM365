import { Injectable } from '@nestjs/common';
import { PrismaService } from '../../core/database/prisma.service';

@Injectable()
export class UsersService {
  constructor(private readonly prisma: PrismaService) {}

  findAll() {
    return this.prisma.user.findMany({
      select: {
        id: true, email: true, name: true, role: true, isActive: true, createdAt: true,
        microsoftConnection: {
          select: { email: true, status: true, lastSyncAt: true },
        },
      },
    });
  }

  findOne(id: string) {
    return this.prisma.user.findUniqueOrThrow({
      where: { id },
      select: {
        id: true, email: true, name: true, role: true, isActive: true, createdAt: true,
        microsoftConnection: {
          select: { email: true, status: true, syncCalendar: true, syncMail: true, syncDocuments: true, lastSyncAt: true, tokenExpiresAt: true },
        },
      },
    });
  }
}
