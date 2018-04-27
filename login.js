const inbox = require('inbox');
const _ = require('underscore');
const later = require('later');
const GoogleSpreadsheet = require('google-spreadsheet');
const moment = require('moment');

const accData = {
  "type": "service_account",
  "project_id": "venta-diaria",
  "private_key_id": "6d16a6c40cee24ccb924269617807d857e137d72",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC2CcyaSNfwRiYn\nLiwNVgQgWQKnpFWHYxIAQD0kKuxhwc7ORS8bXEFKL3s3SUAWnnkHhvuDjvip//nU\ndp+ETO1duUKWx4i8uq2WHKvqDyCP+OHVZtN97zo9P/s2twE5Ps8o7+XrJBsQ8mZy\nUccoSPkpOBexCiYSxYQaVhx/4ME21vGDT9a8YKOLsC36cz4LBUK2pmqbAuBjORKf\nzmWk5w3X8zLufa5CSO1dVjshKHkqi/0KLNm/AwmQW9KBLhPBrvfTHJmn1K52ZXcg\nq+nEzuYZfL+fElGGhCTKcsU42lXrWKYjRYyUJ8XhKcsLsChhPX43D/bjj45H3i1Q\nNJ0mo56zAgMBAAECggEABZ1SbNpyTpGF+7oLgch1YsoGsET786u1PRdoy1oc5ab5\n9JgNq+UHM9KNv6wdXeqQBmYTZoCyYAvqbgh6mF9cXYJT8MfsCmYiPyRYk7ouYZe7\nihQcfHXfG08t4FzPR4S4aaV8jJQV1vVNqO5SG4akqLDM35kAiPOumgwDNQPURxdX\nZcHdxcDpNYjtPtvbfT1OKUOfW7LLdou5i3bmByM4immgdIybBJ8MdjDUsOIzK2DN\nHVx9lh2UmwTevtLzOIQ/LFNRcX9xYwsakCaA3/P8uLuHPZq/KvTirw26YG0AYSwf\n3MUb0o7ms0vikKNI7TzybD9cPpIsPwhVW/8CbfOBgQKBgQDjrLu2gYQuTpdqECPD\nTFcYRlnt78Wkw7xnUztWqt8ZL/4rl9G303vDWkUAiJ1852VpxKAOgOJ7MnAeKrYf\n6ZTW/w0lLghyKPVZSL4j5lWKySBtrrHtD1q6ISr6vV+8pPKDkynPXdIsRSPjjs4e\nR1lJPR4dWGds6Dkl1bx7hDh31wKBgQDMr5RatJxrpCwAZXFK/GJrzUleW+//rMLn\n2wptHVWsrm/v7awEdk63MgReLZ9pUYIJtc+hCLPMxWrONoHTUhcgwYHj/hN4Gowd\nF/ZskNsLWXKdHAej6XoYbnaN7LKPX9V0hKvgCmwRkDZ4eh2eovEVaoVgOPNs32cB\nhr/uAkIEhQKBgQC9SKCPbUJNlX2A6oYxGkjWn7aogM2a3DjI1oPg3BK7SBFSgNgU\nsriUg3oWpX35mA/STWycYj7pGdfo3K2p/nKGBGoTXSAceTzxy+54vkikJ+7UAYdf\nhYJyeJzY9ZSgq6oMBc+e3Wuc7qaVy+ZFeiAbKbrdvt/NxYutjvMy5Yxk5QKBgHwV\nzQgYCeOviQVMehwNWNUlhF7xuVL0Nsw8G9v+NpwSu8Vl/ixOVHX2mnNFkShVw1GD\nqLVlAysWWyNcI+QqFd9DsCy5MLBU17AjgL5cKo580WCxR2h0+BGrla+AWNdWL58N\nduzBJLaZCIyM6zvqZ+ClzOmCXQAZhuaD/AKb183JAoGBAKafqKGyb0/wpjMBgAAB\njAqO8oQ/JnvtA0Md65/nbGBGu64OfKchVNVyK0FryW3nZIiU4GnCStMHhbgcFBDI\nY65bufpnx3ZT6pmhaCOsFZRCLEdrWY+EBOs0X1pvNs9Yw6n5IiTJ9jXc0/p0G1zN\nOaJVWkLSburAGkazZgbEPYdZ\n-----END PRIVATE KEY-----\n",
  "client_email": "editor@venta-diaria.iam.gserviceaccount.com",
  "client_id": "110513764295973423018",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://accounts.google.com/o/oauth2/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/editor%40venta-diaria.iam.gserviceaccount.com"
};

const codes = {
  'Tambo Patio Bellavista Spa': 'TPB',
  'Panko Patio Bellavista Spa': 'PPB',
  'La Criolla Spa': 'LC',
  'Panko Lastarria Spa': 'PL',
  'Soc De Inversiones Y Gastronomica Limo Sa': 'TL',
  'Zambo Bellavista Spa.': 'ZB',
  'Sociedad Gastronomica La Nacional Spa': 'LN'
};

const codeOrder = ['TPB', 'PPB', 'LC', 'PL', 'TL', 'ZB', 'LNLT', 'LNLD'];

const client = inbox.createConnection(false, "imap.gmail.com", {
  secureConnection: true,
  auth:{
    user: "tambobotpagos@gmail.com",
    pass: "holaenfermera2"
  }
});

later.date.UTC();

const monThroughThu = later.parse.text('at 10:00 am on Fri');
const weekend = later.parse.text('at 10:00 am on Mon');
const testSched = later.parse.text('every 5 seconds');

const connectAndWriteSheet = function (sheetName) {
  const logErr = function (type, err) {
    console.error(`${type} error para "${sheetName}":`);
    console.error(err);
  };

  const streamToString = function (stream) {
    const parsingStream = new Promise(function (resolve, reject) {
      var str = '';
      stream.on('data', function (chunk) {
        str += chunk;
      });
      stream.on('end', function () {
        resolve(str);
      });

      stream.on('error', function (err) {
        logErr('stream', err);
      });
    });

    return parsingStream;
  };

  const client = inbox.createConnection(false, "imap.gmail.com", {
    secureConnection: true,
    auth:{
      user: "tambobotpagos@gmail.com",
      pass: "holaenfermera2"
    }
  });
  client.on('connect', function () {
    console.log('connection successful.');

    const doc = new GoogleSpreadsheet('1jcDYf-He8R02i10Y89QypawlSjO1pxveT2W9vSGEPyo');

    const sales = {};
    const nacional = {};
    client.openMailbox('INBOX', function (err, info) {
      if (err) {
        return logErr('mailbox', err);
      }
      client.listMessages(-100, function (err, messages) {
        if (err) {
          return logErr('list', err);
        }

        const gettingAllMails = [];
        _.forEach(messages, function (message) {
          if (!message.title.includes('venta de')) {
            return;
          }

          const stream = client.createMessageStream(message.UID);
          const gettingContent = streamToString(stream)

          .then(function (res) {
            const code = codes[message.title.match(/(?:Fwd: )?(.*?) venta(:?.*?)/)[1]];
            const day = message.title.match(/venta de (.*?) /)[1];
            const amount = parseInt(res.match(/Resumen[\n\r][\n\r]Dia (?:.*?)[\n\r][\n\r]Venta (.*?)[\n\r][\n\r]/)[1].replace(/\./g, ''), 10);

            if (code === 'LN') {
              if (nacional[day]) {
                nacional[day].push(amount);
              } else {
                nacional[day] = [amount];
              }
            } else {
              if (sales[code]) {
                sales[code][day] = amount;
              } else {
                sales[code] = {};
                sales[code][day] = amount;
              }
            }
          })

          .catch(function (err) {
            throw err;
          });

          gettingAllMails.push(gettingContent);
        });

        Promise.all(gettingAllMails)

        .then(function () {
          sales['LNLT'] = {};
          sales['LNLD'] = {};
          _.forEach(nacional, function (amounts, day) {
            sales['LNLT'][day] = _.min(amounts);
            sales['LNLD'][day] = _.max(amounts);
          });

          const weeklySales = {};
          _.forEach(sales, function (rest, key) {
            weeklySales[key] = _.reduce(rest, function (memo, amount) { return memo + amount; });
          });

          doc.useServiceAccountAuth(accData, function () {
            doc.addWorksheet({
              title: sheetName,
              rowCount: 1,
              colCount: 2,
              headers: ['Local', 'Venta']
            }, function (err, sheet) {
              if (err) {
                return logErr('addWorksheet', err);
              }

              const row = {};
              const addRow = function (row) {
                const addingRow = new Promise(function (resolve, reject) {
                  sheet.addRow(row, function (err, addedRow) {
                    if (err) {
                      logErr('addRow', err);
                    }

                    resolve(addedRow);
                  });
                });

                return addingRow;
              };

              const deleteMessage = function (uid) {
                const deletingMessage = new Promise(function (resolve, reject) {
                  client.deleteMessage(uid, function (err) {
                    if (err) {
                      logErr('delete', err);
                    }

                    resolve(uid);
                  });
                });

                return deletingMessage;
              };

              let promiseChain = Promise.resolve();

              const orderedData = _.map(codeOrder, function (code) {
                return {
                  Local: code,
                  Venta: weeklySales[code]
                };
              });

              console.log('DATA: ', orderedData);

              _.forEach(orderedData, function (row) {
                promiseChain = promiseChain.then(function () {
                  return addRow(row);
                });
              });

              promiseChain.then(function () {
                console.log(`Filas agregadas para "${sheetName}"`);
                console.log('Borrando correos...');
                let deletingMessages = Promise.resolve();
                _.forEach(messages, function (message) {
                  deletingMessages = deletingMessages.then(function () {
                    return deleteMessage(message.UID);
                  });
                });

                deletingMessages.then(function () {
                  console.log('Correos borrados');
                  client.close();
                });
              });
            });
          });
        })

        .catch(function (err) {
          logErr('parse', err);
          client.close();
        });
      });
    });
  });

  client.connect();
};

later.setInterval(function () {
  connectAndWriteSheet(`L-J ${moment.utc().subtract(3, 'h').format('DD/MM/YYYY')}`)
}, monThroughThu);

later.setInterval(function () {
  connectAndWriteSheet(`FDS ${moment.utc().subtract(3, 'h').format('DD/MM/YYYY')}`)
}, weekend);

// later.setTimeout(function () {
//   connectAndWriteSheet(`L-J ${moment.utc().subtract(3, 'h').format('DD/MM/YYYY')}`)
// }, testSched);
