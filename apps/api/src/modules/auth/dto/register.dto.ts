import { IsEmail, IsEnum, IsOptional, IsString, MinLength } from 'class-validator';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';

export class RegisterDto {
  @ApiProperty()
  @IsEmail()
  email: string;

  @ApiProperty()
  @IsString()
  name: string;

  @ApiProperty()
  @IsString()
  @MinLength(8)
  password: string;

  @ApiPropertyOptional({ enum: ['ADMIN', 'CLIENT'] })
  @IsOptional()
  @IsEnum(['ADMIN', 'CLIENT'])
  role?: 'ADMIN' | 'CLIENT';
}
