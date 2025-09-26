import { ConfidentialClientApplication, AuthenticationResult } from '@azure/msal-node';
import * as fs from 'fs';
import * as path from 'path';
import { log } from '../utils/logger';

export class TokenManager {
  private cca: ConfidentialClientApplication;
  private cachedTokenResponse: AuthenticationResult | null = null;
  private readonly TOKEN_FILE: string;
  private readonly graphScopes = [
    'Notes.ReadWrite',
    'User.Read',
  ];

  constructor(tenantId: string, clientId: string) {
    log('Initializing TokenManager...');
    
    // Use __dirname to get the directory where this file is located, then go up to project root
    const projectRoot = path.resolve(__dirname, '..', '..');
    this.TOKEN_FILE = path.join(projectRoot, '.auth-token');
    log(`Token file path: ${this.TOKEN_FILE}`);
    
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET || 'not-needed-for-public-client',
      },
    };

    log(`MSAL Authority: ${msalConfig.auth.authority}`);
    
    try {
      this.cca = new ConfidentialClientApplication(msalConfig);
      log('TokenManager initialized successfully');
    } catch (error) {
      log(`ERROR creating MSAL client: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw error;
    }
  }

  async getAccessToken(): Promise<string> {
    log('=== getAccessToken called ===');
    
    try {
      // Try to read pre-saved token from file
      log(`Looking for token file: ${this.TOKEN_FILE}`);
      
      if (fs.existsSync(this.TOKEN_FILE)) {
        log('Token file found, reading...');
        const tokenData = JSON.parse(fs.readFileSync(this.TOKEN_FILE, 'utf8'));
        
        const expiresOn = new Date(tokenData.expiresOn);
        const now = new Date();
        
        log(`Token expires: ${expiresOn}`);
        log(`Current time: ${now}`);
        
        if (expiresOn > now) {
          log('Using valid pre-saved token');
          return tokenData.accessToken;
        } else {
          log('Pre-saved token has expired');
          throw new Error('Authentication token has expired. Please run: npm run auth');
        }
      } else {
        log('No token file found');
        throw new Error('No authentication token found. Please run: npm run auth');
      }
    } catch (error) {
      log(`ERROR in getAccessToken: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw new Error(`Authentication required: ${error instanceof Error ? error.message : 'Unknown error'}. Please run: npm run auth`);
    }
  }

  // Method to check if we have a valid token without throwing
  hasValidToken(): boolean {
    try {
      if (fs.existsSync(this.TOKEN_FILE)) {
        const tokenData = JSON.parse(fs.readFileSync(this.TOKEN_FILE, 'utf8'));
        const expiresOn = new Date(tokenData.expiresOn);
        const now = new Date();
        return expiresOn > now;
      }
      return false;
    } catch {
      return false;
    }
  }

  // Get token expiration info
  getTokenInfo(): { expiresOn?: Date, username?: string, valid: boolean } {
    try {
      if (fs.existsSync(this.TOKEN_FILE)) {
        const tokenData = JSON.parse(fs.readFileSync(this.TOKEN_FILE, 'utf8'));
        const expiresOn = new Date(tokenData.expiresOn);
        const now = new Date();
        
        return {
          expiresOn,
          username: tokenData.account?.username,
          valid: expiresOn > now
        };
      }
      return { valid: false };
    } catch {
      return { valid: false };
    }
  }
}

