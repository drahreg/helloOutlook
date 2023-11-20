const express = require('express');
const path = require('path');
const https = require('https');
const fs = require('fs');
const app = express();
const port = 3000;
const serverOptions = {
  key: fs.readFileSync('../../cert/localhost.key'),
  cert: fs.readFileSync('../../cert/localhost.crt')
};


/* 
 * This server would check the inputs of the user and request the API
 * and forward the results
 * 
 * At the moment i just return a text based on the inputs of the user to
 * test the Office Addin prototype
 */ 

app.get('/', (req, res) => {
  let email_type = req.query.email_type;
  let gender = req.query.gender;
  let name = req.query.name;

  let result = '';

  result += 'Hallo';

  if ((gender != undefined) && (name != undefined)) {
    result += ' ' + gender + ' ' + name + ',\n';
  } else {
    result += ',\n';
  }

  switch (email_type) {
    case 'meeting_request' :
      result += 'Es handelt sich um eine Terminanfrage';
      break;
    case 'status_update' :
      result += 'Es handelt sich um ein Status Update';
      break;
    case 'product_offer' :
      result += 'Es handelt sich um eine Produkt Angebot';
      break;
    default: 
      break;
  }

  res.set('Access-Control-Allow-Origin', '*')
  res.send(result);
})

app.get('/taskpane.html', function(req, res) {
  res.sendFile(path.join(__dirname, '/taskpane.html'));
});

const server = https.createServer(serverOptions, app);

server.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
