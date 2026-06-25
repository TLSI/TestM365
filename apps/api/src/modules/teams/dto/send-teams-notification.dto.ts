import { IsString, IsUrl, IsOptional, IsArray, ValidateNested } from 'class-validator';
import { Type } from 'class-transformer';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';

class TeamsFact {
  @ApiProperty()
  @IsString()
  name: string;

  @ApiProperty()
  @IsString()
  value: string;
}

export class SendTeamsNotificationDto {
  @ApiProperty({ description: 'Incoming webhook URL from Teams channel' })
  @IsUrl()
  webhookUrl: string;

  @ApiProperty()
  @IsString()
  title: string;

  @ApiPropertyOptional()
  @IsOptional()
  @IsString()
  subtitle?: string;

  @ApiProperty()
  @IsString()
  body: string;

  @ApiPropertyOptional({ type: [TeamsFact] })
  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => TeamsFact)
  facts?: TeamsFact[];
}
