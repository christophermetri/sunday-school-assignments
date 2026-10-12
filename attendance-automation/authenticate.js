const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const http = require('http');
const url = require('url');

const SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://mail.google.com/'
];

const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

async function authenticate() {
  try {
    // Load client credentials
    if (!fs.existsSync(CREDENTIALS_PATH)) {
      console.error('❌ credentials.json not found!');
      console.log('Download it from: https://console.cloud.google.com/apis/credentials/oauthclient');
      process.exit(1);
    }

    const credentialsContent = fs.readFileSync(CREDENTIALS_PATH, 'utf8');
    const { client_id, client_secret, redirect_uris } = JSON.parse(credentialsContent).installed;

    const oauth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[0]
    );

    // Generate auth URL
    const authUrl = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
      prompt: 'consent'
    });

    console.log('\n🔐 Authentication Required');
    console.log('=====================================');
    console.log('Visit this URL to authorize the app:');
    console.log(authUrl);
    console.log('=====================================\n');

    // Start local server to receive callback
    const server = http.createServer(async (req, res) => {
      const parsedUrl = url.parse(req.url, true);
      const code = parsedUrl.query.code;

      if (code) {
        try {
          const { tokens } = await oauth2Client.getToken(code);
          
          // Save tokens
          fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
          console.log('✅ Authentication successful!');
          console.log('✅ Tokens saved to token.json');
          
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end('<h1>✅ Authentication successful! You can close this window.</h1>');
          
          server.close();
          process.exit(0);
        } catch (error) {
          console.error('❌ Error getting tokens:', error.message);
          res.writeHead(400);
          res.end('❌ Authentication failed');
          server.close();
          process.exit(1);
        }
      } else {
        res.writeHead(400);
        res.end('❌ No authorization code received');
        server.close();
        process.exit(1);
      }
    });

    server.listen(3000, () => {
      console.log('📌 Local server listening on http://localhost:3000');
      console.log('(You will be redirected here after authorizing)\n');
    });

  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

authenticate();
