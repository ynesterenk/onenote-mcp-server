import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from '../auth/tokenManager';
import { log } from './logger';

export function createGraphClient(tokenManager: TokenManager): Client {
  log('Creating Microsoft Graph client...');
  
  const graphClient = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        return await tokenManager.getAccessToken();
      }
    }
  });
  
  log('Microsoft Graph client created successfully');
  return graphClient;
}

