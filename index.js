const APP_ID = `XXXXXXXXXXXXXXXXXXXXX`; // The client_id of the microsoft app registration
const APP_SECERET = `XXXXXXXXXXXXXXXXX`; // The client_secret of the microsoft app registration
const GROUP_ID = `XXXXXXXXXXXXXXXXXXXXX`; // The group ID to invite people into

const SENDGRID_KEY = `XXXXXXXXXXXXXXXXXX`; // The API key for sendgrid to send the email
const TEMPLATE_ID = `XXXXXXXXXXXXXXXXXXXXXXXX`; // The ID of the dynamic template in sendgrid for the email

const MONGODB_URI = `XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX`; // The cosmos db mongo connection string
const MONGODB_DBNAME = `XXXXXXXXXXXX`; // The DB name
const MONGODB_COLLECTION = `XXXXXXXXXXXXX`; // The collection name

const TOKEN_ENDPOINT =`https://login.microsoftonline.com/XXXXXXXXXXXXXXXXXXXXX/oauth2/v2.0/token`; // Add the tennant ID instead of the Xs
const MS_GRAPH_SCOPE = `https://graph.microsoft.com/.default`;
const INVITE_URL = `https://graph.microsoft.com/v1.0/invitations`;

const axios = require('axios');
const assert = require('assert');
const qs = require('qs');
const sgMail = require('@sendgrid/mail');
const MongoClient = require('mongodb').MongoClient;


sgMail.setApiKey(SENDGRID_KEY);

const postData = {
  client_id: APP_ID,
  scope: MS_GRAPH_SCOPE,
  client_secret: APP_SECERET,
  grant_type: 'client_credentials'
};

axios.defaults.headers.post['Content-Type'] =
  'application/x-www-form-urlencoded';

const saveToDB = (name, email, phone, agree_to_emails, terms, profession, wix_id, wix_created_at, wix_submission_time, userId, invitationLink) => {
    MongoClient.connect(MONGODB_URI, function(err, client) {
        assert.equal(null, err);
        console.log("Connected successfully to server");
        
        const db = client.db(MONGODB_DBNAME);
        const collection = db.collection(MONGODB_COLLECTION);

        const data = {
            name,
            email,
            phone,
            agree_to_emails,
            terms,
            profession,
            wixId: wix_id,
            wixCreatedAt: wix_created_at,
            wixSubmissionTime: wix_submission_time,
            userId,
            invitationLink,
            createdAt: (new Date).getTime(),
        };

        collection.insertOne(data, function(err, res) {
            if (err) throw err;
            console.log("1 document inserted");
        });
        client.close();
    });
}

const getToken = async () => {
    const resp = await axios.post(TOKEN_ENDPOINT, qs.stringify(postData));
    return resp.data.access_token;
}

const inviteToOrg = async (email, name, token) => {
    const data = {
        "invitedUserEmailAddress": email,
        "inviteRedirectUrl": "https://teams.microsoft.com",
        "sendInvitationMessage": false,
        "invitedUserType": "Member",
        "invitedUserDisplayName": name,
    };
    const resp = await axios.post(INVITE_URL, data, {headers: {Authorization: `Bearer ${token}`}});
    return resp.data;
}

const addToTeam = async (userId, token) => {
    const directoryObject = {
        "@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}`,
    };

    const resp = await axios.post(`https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/members/$ref`, directoryObject, {headers: {Authorization: `Bearer ${token}`}});
    return resp;
}

const sendEmail = (email, name, invitationLink) => {
    const msg = {
        to: email,
        from: {
            email: 'hackathon@in.dev',
            name: 'Hackathon Registration'
        },
        templateId: TEMPLATE_ID,
        dynamic_template_data: {
            name,
            invitationLink,
        }
    };
    sgMail.send(msg);
}

module.exports = async function (context, req) {
    const { name, email, phone, agree_to_emails, terms, profession, wix_id, wix_created_at, wix_submission_time } = req.body;

    const token = await getToken();

    try {
        const { invitedUser, inviteRedeemUrl } = await inviteToOrg(email, name, token);
        await addToTeam(invitedUser.id, token);
        sendEmail(email, name, inviteRedeemUrl);
        saveToDB(name, email, phone, agree_to_emails, terms, profession, wix_id, wix_created_at, wix_submission_time, invitedUser.id, inviteRedeemUrl);
    } catch(error) {
        saveToDB(name, email, phone, agree_to_emails, terms, profession, wix_id, wix_created_at, wix_submission_time, "NO_USER_ID", "NO_REDEEM_URL");
    }

    context.res = {
        body: JSON.stringify({success: true})
    };
};
