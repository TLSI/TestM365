import { IsArray, IsEmail, IsOptional, IsString } from 'class-validator';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';

export class SendEmailDto {
  @ApiProperty({ type: [String] })
  @IsArray()
  @IsEmail({}, { each: true })
  to: string[];

  @ApiPropertyOptional({ type: [String] })
  @IsOptional()
  @IsArray()
  cc?: string[];

  @ApiProperty()
  @IsString()
  subject: string;

  @ApiProperty({ description: 'HTML body' })
  @IsString()
  body: string;

  @ApiPropertyOptional({ description: 'Link email to a client record' })
  @IsOptional()
  @IsString()
  clientId?: string;
}
