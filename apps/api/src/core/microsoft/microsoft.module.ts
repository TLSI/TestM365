import { Global, Module } from '@nestjs/common';
import { MicrosoftGraphService } from './microsoft-graph.service';

@Global()
@Module({
  providers: [MicrosoftGraphService],
  exports: [MicrosoftGraphService],
})
export class MicrosoftModule {}
