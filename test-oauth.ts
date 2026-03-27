import { google } from 'googleapis';

const oauth2Client = new google.auth.OAuth2(
  undefined,
  undefined,
  'http://localhost:3000/auth/callback'
);

const url = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  prompt: 'consent',
  scope: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
  ],
  state: 'http://localhost:3000/auth/callback'
});

console.log(url);
