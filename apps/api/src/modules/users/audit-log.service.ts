import { Injectable } from '@nestjs/common';
import { PrismaService } from '../../core/database/prisma.service';

interface AuditLogEntry {
  userId?: string;
  action: string;
  resource: string;
  resourceId?: string;
  details?: Record<string, unknown>;
  ip?: string;
}

@Injectable()
export class AuditLogService {
  constructor(private readonly prisma: PrismaService) {}

  async log(entry: AuditLogEntry) {
    const { userId, ...rest } = entry;
    return this.prisma.auditLog.create({
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      data: { ...rest, ...(userId ? { userId } : {}) } as any,
    });
  }

  async findAll(limit = 100) {
    return this.prisma.auditLog.findMany({
      orderBy: { createdAt: 'desc' },
      take: limit,
      include: { user: { select: { email: true, name: true } } },
    });
  }
}
