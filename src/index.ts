import 'dotenv/config';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

const credential = new ClientSecretCredential(
  process.env.TENANT_ID!,
  process.env.CLIENT_ID!,
  process.env.CLIENT_SECRET!
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default'],
});

const graph = Client.initWithMiddleware({ middleware: authProvider });

(async () => {
  const user = await graph.api('/me').get();
  console.log('Hello', user.displayName);

  const messages = await graph
    .api('/me/messages')
    .top(5)
    .select('subject,from')
    .get();
  messages.value.forEach((m: any) => console.log('Mail:', m.subject));

  const root = await graph.api('/me/drive/root/children').get();
  console.log('Root files count', root.value.length);
})();
