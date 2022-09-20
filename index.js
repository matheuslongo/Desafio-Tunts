const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

const spreadsheetId =  '1eIb13WwR0RSdnBb0RtyJAAK15NpWIx3KqQOXHMFGB8I'

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content), listStudents);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Get all informations on the datasheet
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function listStudents(auth) {
  const sheets = google.sheets({version: 'v4', auth});
  sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'A2:H',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const rows = res.data.values;
    if (rows.length) {
      const avaliatedStudents = avaliateStudents(rows)
      updateDatasheet(avaliatedStudents,sheets)
    } else {
      console.log('No data found.');
    }
  });
}

//Update datasheet at googleapi
const updateDatasheet = (avaliatedStudents, sheets) => {
  const resource = {
  values:avaliatedStudents,
};
  sheets.spreadsheets.values.update({
  spreadsheetId,
  range: 'G4:H',
  valueInputOption:'RAW',
  resource,
}, (err, result) => {
  if (err) {
    console.log(err);
  } else {
    console.log('%d cells updated.', result.data.updatedCells);
  }
})
}

const avaliateStudents = (rows) => {
  const [classRow] = rows[0]
  const classes = classRow.split(" ").pop()
  rows.shift()
  rows.shift()
  console.log("Total de aulas:", classes)

  const result =  rows.map((row) => {
    const [situation,finalNote] = verifySituation(classes, row)
    row[6] = situation
    row[7] = finalNote
    console.log(row)
    return [situation,finalNote]
  });

  return result
}

const verifySituation = (classes, [ , , frequency, P1, P2, P3]) => {
  if ((classes / 4) < frequency) return ['Reprovado por Falta',0]
  const note = Math.round((parseInt(P1) + parseInt(P2) + parseInt(P3)) / 3)
  if(note>=70) return ['Aprovado',0]
  if(note>=50) return ['Exame Final',verifyFinalNote(note)]
  return ['Reprovado por Nota',0]
}

const verifyFinalNote  = (note) => 100 - note