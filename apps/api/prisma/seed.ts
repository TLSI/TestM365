import { PrismaClient } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function main() {
  const hash = await bcrypt.hash('Admin1234!', 10);

  const admin = await prisma.user.upsert({
    where: { email: 'admin@testm365.local' },
    update: {},
    create: {
      email: 'admin@testm365.local',
      name: 'Admin TestM365',
      password: hash,
      role: 'ADMIN',
    },
  });

  const acme = await prisma.client.upsert({
    where: { email: 'acme@example.com' },
    update: {},
    create: {
      name: 'ACME SL',
      email: 'acme@example.com',
      company: 'ACME SL',
    },
  });

  const client = await prisma.user.upsert({
    where: { email: 'cliente@testm365.local' },
    update: { clientId: acme.id },
    create: {
      email: 'cliente@testm365.local',
      name: 'Cliente Demo',
      password: hash,
      role: 'CLIENT',
      clientId: acme.id,
    },
  });

  console.log(`Seed complete. Admin: ${admin.email} | Client: ${client.email}`);
  console.log('Password for both: Admin1234!');
}

main()
  .catch(console.error)
  .finally(() => prisma.$disconnect());
