const json2xls = require('json2xls');
const request = require('request');
const fs = require('fs');
const moment = require('moment');
const readline = require('readline');
const util = require('util');
const exec = util.promisify(require('child_process').exec);
const { google } = require('googleapis');
const log4js = require('log4js');
const logger = log4js.getLogger();
const config = require('./config.json');
log4js.configure({
  appenders: {
    alerts: { type: '@log4js-node/slack', token: config.slackToken, channel_id: 'ut-exporter', username: 'Log4js' }
  },
  categories: {
    default: { appenders: ['alerts'], level: 'debug' }
  }
});
const XLSX = require('xlsx');
const sleep = require('sleep');
const fse = require('fs-extra');
const path = require('path');


const defaultUrl = 'https://api.hubapi.com/deals/v1/deal/paged?hapikey\=0fcca918-bb3f-48a8-9f9b-ecbbae3336dd&limit=250';
let urlWithOffset = '';
let dealIds = [];
let deals = [];
let owners = [];
let dealsXlsxJson = [];
let dealIndex = 0;
let offset;
const dateOfUpload = moment().format('DD-MM-YYYY');
const FILES_DIR = './files/';
const PODIO_ITEMS_EXPORT_PATH = `${FILES_DIR}PodioItemsExport.xlsx`;
const HUBSPOT_DEALS_EXPORT_PATH = `${FILES_DIR}HubspotDealsExport.xlsx`;
const HUBSPOT_PODIO_EXPORT_PATH = `${FILES_DIR}HubspotPodioExport.xlsx`;
const HUBSPOT_PODIO_EXPORT_DATE_PATH = `${FILES_DIR}HubspotPodioExport${dateOfUpload}.xlsx`;

executeRubyAndContinueProcess();

async function executeRubyAndContinueProcess() {
  logger.info("Podio items are being fetched...");
  console.log('Podio items are being fetched...');
  const { error, stdout, stderr } = await exec('ruby fetchPodio.rb');
  if (stdout) {
    console.log(stdout);
    logger.info(stdout);
  }
  if (error) {
    logger.error('Error:', stderr);
    console.log('Error:', stderr);
    throw error;
  }
  getDealIds(defaultUrl, function () {
    console.log('dealIds', dealIds.length);
    logger.info('dealIds', dealIds.length);

    getDeals(function () {
      // saveDealsAsXlsx();
      getOwners();
      // mergeWorkbooks();
      // uploadToDrive();
    })
  });
};


function getDealIds(url, callback) {
  request(url, function (error, response, body) {
    if (error) {
      console.log(error);
      logger.error(error);
      getDealIds(urlWithOffset, callback);
      sleep.sleep(5);
    };
    try {


      const responseObj = JSON.parse(body);

      let dealId = '';
      for (let i = 0; i < responseObj.deals.length; i++) {
        dealId = responseObj.deals[i].dealId;
        dealIds.push(dealId);
      }

      if (responseObj.hasMore) {
        offset = responseObj.offset;
        urlWithOffset = `https://api.hubapi.com/deals/v1/deal/paged?hapikey\=0fcca918-bb3f-48a8-9f9b-ecbbae3336dd&limit=250&offset=${offset}`;

        getDealIds(urlWithOffset, callback);
      } else {
        callback();
      }

    } catch (e) {
      console.log(e);
      logger.error(e);
      getDealIds(urlWithOffset, callback);
      sleep.sleep(5);
    }

  });
}

function getDeals(callback) {
  let currDealId = dealIds[dealIndex];
  // console.log('dealID', currDealId);

  request(`https://api.hubapi.com/deals/v1/deal/${currDealId}?hapikey\=0fcca918-bb3f-48a8-9f9b-ecbbae3336dd`, function (error, response, body) {
    if (error) {
      console.log(error);
      logger.error(error);
      console.log(`Error Deal Id: `, dealIds[dealIndex]);
      sleep.sleep(5);
      getDeals(callback);
    };

    try {
      const responseJson = JSON.parse(body);
      deals.push(responseJson);

      dealIndex++;
      if (dealIndex < dealIds.length) {
        // if (dealIndex < 20) {
        if (dealIndex % 10 == 0) {
          console.log(`Deal Id: ${responseJson.dealId}, Index: ${dealIndex}`);
          logger.info(`Deal Id: ${responseJson.dealId}, Index: ${dealIndex}`);
        }
        // sleep.msleep(100);
        getDeals(callback);
      } else {
        callback();
      }
    } catch (e) {
      console.log(e);
      logger.error(e);
      console.log(`Error Deal Id: `, dealIds[dealIndex]);
      console.log(`Error response body: `, body);
      dealIndex++;
      sleep.sleep(5);
      getDeals(callback);
    }
  });
}
function getOwners() {
  request(`https://api.hubapi.com/owners/v2/owners/?hapikey\=0fcca918-bb3f-48a8-9f9b-ecbbae3336dd`, function (error, response, body) {
    if (error) {
      console.log(error);
      logger.error(error);
      sleep.sleep(5);
      getOwners()
    };

    try {
      const responseJson = JSON.parse(body);
      owners = responseJson;
      saveDealsAsXlsx();

    } catch (e) {
      console.log(e);
      logger.error(e);
      sleep.sleep(5);
      getOwners();
    }
  });
}

function saveDealsAsXlsx() {

  //replace forEach with loop
  // for (const deal of deals) {
  deals.forEach(deal => {
    const dealId = deal.dealId;
    const properties = deal.properties;

    if (!properties && !dealId) {
      // logger.info(`No deal id or properties `, deal)
      console.log(`No deal id or properties `)
      return;
    }

    // const pipeline = properties.pipeline ? properties.pipeline.value : ''; has default value
    // const notes_last_activity_date = properties.notes_last_activity_date ? moment(parseInt(properties.notes_last_activity_date.value)).format('YYYY-MM-DD HH:mm') : ''; same as last contacted
    const m_de3_booket_dato = (!!properties.m_de3_booket_dato && !!properties.m_de3_booket_dato.value) ? moment(parseInt(properties.m_de3_booket_dato.value)).format('YYYY-MM-DD HH:mm') : '';
    const hs_lastmodifieddate = (!!properties.hs_lastmodifieddate && !!properties.hs_lastmodifieddate.value) ? moment(parseInt(properties.hs_lastmodifieddate.value)).format('YYYY-MM-DD HH:mm') : '';
    const closedate = (!!properties.closedate && !!properties.closedate.value) ? moment(parseInt(properties.closedate.value)).format('YYYY-MM-DD HH:mm') : '';
    const createdate = (!!properties.createdate && !!properties.createdate.value) ? moment(parseInt(properties.createdate.value)).format('YYYY-MM-DD HH:mm') : '';
    const i_have_attached_the_accepted_sla_terms = (!!properties.i_have_attached_the_accepted_sla_terms && !!properties.i_have_attached_the_accepted_sla_terms.value) ? properties.i_have_attached_the_accepted_sla_terms.value : '';
    const closed_won_reason = (!!properties.closed_won_reason && !!properties.closed_won_reason.value) ? properties.closed_won_reason.value : '';
    const credit_price = (!!properties.credit_price && !!properties.credit_price.value) ? properties.credit_price.value : '';
    const level = (!!properties.level && !!properties.level.value) ? properties.level.value : '';
    const dealtype = (!!properties.dealtype && !!properties.dealtype.value) ? properties.dealtype.value : '';
    const salestype = (!!properties.salestype && !!properties.salestype.value) ? properties.salestype.value : '';
    const num_contacted_notes = (!!properties.num_contacted_notes && !!properties.num_contacted_notes.value) ? properties.num_contacted_notes.value : '';
    const num_notes = (!!properties.num_notes && !!properties.num_notes.value) ? properties.num_notes.value : '';
    const closed_lost_reason = (!!properties.closed_lost_reason && !!properties.closed_lost_reason) ? properties.closed_lost_reason.value : '';
    const product = (!!properties.product && !!properties.product.value) ? properties.product.value : '';
    const hubspot_owner_id = (!!properties.hubspot_owner_id && !!properties.hubspot_owner_id.value) ? properties.hubspot_owner_id.value : '';
    const notes_next_activity_date = (!!properties.notes_next_activity_date && !!properties.notes_next_activity_date.value) ? moment(parseInt(properties.notes_next_activity_date.value)).format('YYYY-MM-DD HH:mm') : '';
    const tilbud_sendt_dato = (!!properties.tilbud_send && !!properties.tilbud_send.value) ? moment(parseInt(properties.tilbud_sendt_dato.value)).format('YYYY-MM-DD HH:mm') : '';
    const hubspot_owner_assigneddate = (!!properties.hubspot_owner_assigneddate && !!properties.hubspot_owner_assigneddate.value) ? moment(parseInt(properties.hubspot_owner_assigneddate.value)).format('YYYY-MM-DD HH:mm') : '';
    const strategiskeudfordringer = (!!properties.strategiskeudfordringer && !!properties.strategiskeudfordringer.value) ? properties.strategiskeudfordringer.value : '';
    const dealstage = (!!properties.dealstage && !!properties.dealstage.value) ? properties.dealstage.value : '';
    const num_associated_contacts = (!!properties.num_associated_contacts && !!properties.num_associated_contacts.value) ? properties.num_associated_contacts.value : '';
    const hs_analytics_source_data_1 = (!!properties.hs_analytics_source_data_1 && !!properties.hs_analytics_source_data_1.value) ? properties.hs_analytics_source_data_1.value : '';
    const m_de_booket_dato = (!!properties.m_de_booket_dato && !!properties.m_de_booket_dato.value) ? moment(parseInt(properties.m_de_booket_dato.value)).format('YYYY-MM-DD HH:mm') : '';
    const hs_analytics_source_data_2 = (!!properties.hs_analytics_source_data_2 && !!properties.hs_analytics_source_data_2.value) ? properties.hs_analytics_source_data_2.value : '';
    const scope = (!!properties.scope && !!properties.scope.value) ? properties.scope.value : '';
    const notes_last_contacted = (!!properties.notes_last_contacted && !!properties.notes_last_contacted.value) ? moment(parseInt(properties.notes_last_contacted.value)).format('YYYY-MM-DD HH:mm') : '';
    const hubspot_team_id = (!!properties.hubspot_team_id && !!properties.hubspot_team_id.value) ? properties.hubspot_team_id.value : '';
    const won = (!!properties.won && !!properties.won.value) ? properties.won.value : '';
    const country = (!!properties.country && !!properties.country.value) ? properties.country.value : '';
    const dealname = (!!properties.dealname && !!properties.dealname.value) ? properties.dealname.value : '';
    const saleschannel = (!!properties.saleschannel && !!properties.saleschannel.value) ? properties.saleschannel.value : '';
    const amount_dropdown = (!!properties.amount_dropdown && !!properties.amount_dropdown.value) ? properties.amount_dropdown.value : '';
    const amount = (!!properties.amount && !!properties.amount.value) ? properties.amount.value : '';
    const solgt_dato = (!!properties.solgt_dato && !!properties.solgt_dato.value) ? moment(parseInt(properties.solgt_dato.value)).format('YYYY-MM-DD HH:mm') : '';
    const credits = (!!properties.credits && !!properties.credits.value) ? properties.credits.value : '';
    let creditsInteger = '';
    if (credits) {
      creditsInteger = parseInt(credits);
    }

    let ownerEmail;
    for (const owner of owners) {
      if (owner.ownerId == hubspot_owner_id) {
        ownerEmail = owner.email;
      }
    }
    if (!ownerEmail) {
      ownerEmail = 'Deactivated User'
    }

    // if (!!hubspot_owner_id) {
    //   owners.forEach(owner => {
    //     if (owner.ownerId == hubspot_owner_id) {
    //       ownerEmail = owner.email;
    //       return;
    //     }
    //   });
    // }
    // if (!hubspot_owner_id) {
    //   ownerEmail = 'Deactivated User'
    // }

    const source = (!!properties.source && !!properties.source) ? properties.source.value : '';
    const description = (!!properties.description && !!properties.description) ? properties.description.value : '';
    const m_de2_booket_dato = (!!properties.m_de2_booket_dato && !!properties.m_de2_booket_dato) ? moment(parseInt(properties.m_de2_booket_dato.value)).format('YYYY-MM-DD HH:mm') : '';
    const associatedCompanyId = (deal.associations.associatedCompanyIds[0] && !!deal.associations.associatedCompanyIds[0]) ? deal.associations.associatedCompanyIds[0] : '';

    dealsXlsxJson.push({
      "Deal Id": dealId,
      "Date of 3 meeting booked": m_de3_booket_dato,
      "Closed Won Reason": closed_won_reason,
      "I have attached the accepted SLA terms": i_have_attached_the_accepted_sla_terms,
      "Last Modified Date": hs_lastmodifieddate,
      "Pipeline": 'Sales Pipeline', //default value
      "Credit price": credit_price,
      "Level": level,
      "Close Date": closedate,
      "Deal Type": dealtype,
      "Type of sale": salestype,
      "Number of times contacted": num_contacted_notes,
      "Number of Sales Activities": num_notes,
      "Auto-renewal": '', //not important
      "Original Source Type": '', //not important
      "CI_analyst": '', //not important
      "Create Date": createdate,
      "Closed Lost Reason": closed_lost_reason,
      "ConsultantLead": '', //field not in response
      "Proposal/accept": '', //not important
      "Co-create": '', //field not in response
      "Product type": product,
      // "Deal owner Id": hubspot_owner_id,
      "Deal owner": ownerEmail,
      "Last Activity Date": notes_last_contacted, // same as last contacted
      "Next Activity Date": notes_next_activity_date,
      "Offer sent date": tilbud_sendt_dato,
      "Owner Assigned Date": hubspot_owner_assigneddate,
      "Strategic challenges": strategiskeudfordringer,
      "Deal Stage": dealstage,
      "Number of Contacts": num_associated_contacts, // field not with name in response
      "Original Source Data 1": hs_analytics_source_data_1,
      "Date of meeting booked": m_de_booket_dato,
      "Original Source Data 2": hs_analytics_source_data_2,
      "Scope": scope,
      "Last Contacted": notes_last_contacted,
      "HubSpot Team": hubspot_team_id,
      "Won": won,
      "Country": country,
      "Deal Name": dealname,
      "Sales channel": saleschannel,
      "Amount dropdown": amount_dropdown,
      "Amount": amount,
      "Date sold": solgt_dato,
      "Credits": creditsInteger,
      "Source of meeting booked": source,
      "Deal Description": description,
      "Date of 2 meeting booked": m_de2_booket_dato,
      "Associated Company Id": associatedCompanyId,
      "Associated Contacts": num_associated_contacts
    })
  });
  // }
  // } catch (e) {
  //   console.log(e);

  // }
  writeHubspot();
}

function writeHubspot() {
  const xls = json2xls(dealsXlsxJson);
  fs.writeFileSync(HUBSPOT_DEALS_EXPORT_PATH, xls, 'binary');
  console.log(`${deals.length} deals records saved to HubspotDealsExport.xlsx to "files" folder`);
  logger.info(`${deals.length} deals records saved to HubspotDealsExport.xlsx to "files" folder`);

  mergeWorkbooks();
}

function mergeWorkbooks() {
  const outWb = XLSX.utils.book_new();

  const hubspotWb = XLSX.readFile(HUBSPOT_DEALS_EXPORT_PATH);
  const podioWb = XLSX.readFile(PODIO_ITEMS_EXPORT_PATH);

  const hubspotWs = hubspotWb.Sheets['Sheet 1'];
  const podioWs = podioWb.Sheets['Deliverances'];

  XLSX.utils.book_append_sheet(outWb, hubspotWs, 'HubspotDeals');
  XLSX.utils.book_append_sheet(outWb, podioWs, 'PodioItems');

  XLSX.writeFile(outWb, HUBSPOT_PODIO_EXPORT_PATH);
  console.log(`HubspotPodioExport.xlsx saved to "files" folder`);
  logger.info(`HubspotPodioExport.xlsx saved to "files" folder`);

  fse.copySync(path.resolve(__dirname, HUBSPOT_PODIO_EXPORT_PATH), HUBSPOT_PODIO_EXPORT_DATE_PATH);
  console.log(`HubspotPodioExport${dateOfUpload}.xlsx saved to "files" folder`);
  logger.info(`HubspotPodioExport${dateOfUpload}.xlsx saved to "files" folder`);

  uploadToDrive();
}

function uploadToDrive() {
  // If modifying these scopes, delete token.json.
  const SCOPES = ['https://www.googleapis.com/auth/drive'];
  const TOKEN_PATH = 'token.json';

  // Load client secrets from a local file.
  fs.readFile('credentials.json', (err, content) => {
    if (err) {
      logger.error(err);
      return console.log('Error loading client secret file:', err);
    }
    // Authorize a client with credentials, then call the Google Drive API.
    authorize(JSON.parse(content), insertFilesInFolder);
  });

  /**
   * Create an OAuth2 client with the given credentials, and then execute the
   * given callback function.
   * @param {Object} credentials The authorization client credentials.
   * @param {function} callback The callback to call with the authorized client.
   */
  function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
      if (err) {
        logger.error(err);
        return getAccessToken(oAuth2Client, callback);
      }
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
  function getAccessToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
    console.log('Authorize this app by visiting this url:', authUrl);
    logger.info('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, });
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken(code, (err, token) => {
        if (err) {
          logger.error(err);
          return console.error('Error retrieving access token', err);
        }
        oAuth2Client.setCredentials(token);
        // Store the token to disk for later program executions
        fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
          if (err) console.error(err);
          console.log('Token stored to', TOKEN_PATH);
          logger.info('Token stored to', TOKEN_PATH);
        });
        callback(oAuth2Client);
      });
    });
  }

  /**
   * Lists the names and IDs of up to 10 files.
   * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
   */

  function insertFilesInFolder(auth) {
    const drive = google.drive({ version: 'v3', auth });
    const folderId = '1SC_6YcwF_8_GscPg1L5bvSiinr7AKdQ7';
    const fileId = '1Rh4aR5n3mlu-D5yLC4WTcteul31kLjkG';

    // file id to be updated 1gG8DbDSz4jeX4lm9f1Ce1P_nbWN_0dote-kjag2gmW8
    createFile(HUBSPOT_PODIO_EXPORT_DATE_PATH, `HubspotPodioExport${dateOfUpload}.xlsx`);
    updateFile(HUBSPOT_PODIO_EXPORT_PATH, `HubspotPodioExport.xlsx`, fileId);

    function createFile(filePath, name) {
      const fileMetadata = { 'name': name, parents: [folderId] };
      const media = { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: fs.createReadStream(filePath) };
      drive.files.create({
        resource: fileMetadata, media: media, fields: 'id'
      }, function (err, file) {
        if (err) {
          // Handle error
          console.error(err);
          logger.error(err);
        } else {
          console.log(`Created ${name} to Google Drive with File Id: ${file.data.id}`);
          logger.info(`Created ${name} to Google Drive with File Id: ${file.data.id}`);

          if (fs.existsSync(HUBSPOT_PODIO_EXPORT_DATE_PATH)) {
            try {
              fs.unlinkSync(HUBSPOT_PODIO_EXPORT_DATE_PATH);
              logger.info('Files HUBSPOT_PODIO_EXPORT_DATE_PATH successfully deleted from local storage');
              console.log('Files HUBSPOT_PODIO_EXPORT_DATE_PATH successfully deleted from local storage');
            } catch (e) {
              console.log(e);
              logger.info(e);
            }
          }

        }
      });
    }
    function updateFile(filePath, name, fileId) {
      const fileMetadata = {
        'name': name, addParents: [folderId]
      };
      const media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: fs.createReadStream(filePath)
      };
      drive.files.update({
        fileId: fileId, resource: fileMetadata, media: media, fields: 'id'
      }, function (err, file) {
        if (err) {
          // Handle error
          console.error(err);
          logger.error(err);
        } else {
          console.log(`Updated ${name} to Google Drive with File Id: ${file.data.id}`);
          logger.info(`Updated ${name} to Google Drive with File Id: ${file.data.id}`);
        }
      });
    }
  }
}