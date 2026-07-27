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
  @ApiProperty({ description: 'Teams channel URL (deep link from "Get link to channel")' })
  @IsUrl()
  channelUrl: string;

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
