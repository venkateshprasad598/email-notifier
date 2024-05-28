const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const qs = require('qs');

const app = express();
app.use(bodyParser.json());

const clientId = '03195911-8786-406f-a12d-bc76012fe957';
const clientSecret = '89c6f147-94f7-46b6-a424-c1b2f9fcdda6';
const tenantId = 'YOUR_TENANT_ID';
const notificationUrl = 'https://your-server-url/webhook';
const clientState = 'yourClientState';

const getAccessToken = async () => {
    try {
        const response = await axios.post(
            `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            qs.stringify({
                grant_type: 'client_credentials',
                client_id: clientId,
                client_secret: clientSecret,
                scope: 'https://graph.microsoft.com/.default',
            }),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );
        return response.data.access_token;
    } catch (error) {
        console.error('Error getting access token:', error);
        throw error;
    }
};

const createSubscription = async (accessToken) => {
    try {
        const response = await axios.post(
            'https://graph.microsoft.com/v1.0/subscriptions',
            {
                changeType: 'created',
                notificationUrl: notificationUrl,
                resource: 'me/messages',
                expirationDateTime: new Date(new Date().getTime() + 3600 * 1000).toISOString(),
                clientState: clientState,
            },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                },
            }
        );
        return response.data;
    } catch (error) {
        console.error('Error creating subscription:', error);
        throw error;
    }
};

const handleNotifications = async (notifications) => {
    for (const notification of notifications.value) {
        if (notification.clientState === clientState) {
            const messageId = notification.resourceData.id;
            const accessToken = await getAccessToken();

            try {
                const response = await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${messageId}`, {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-Type': 'application/json',
                    },
                });
                const email = response.data;
                // Process email (e.g., save body and attachments)
                console.log('New email received:', email);
                res.sendStatus(202).json({ email })
            } catch (error) {
                console.error('Error fetching new email:', error);
            }
        }
    }
};



app.get('/', async (req, res) => {
    res.status(200).json("Hello Venky")
});

app.post('/webhook', async (req, res) => {
    const { validationToken } = req.query;
    if (validationToken) {
        res.status(200).send(validationToken);
    } else {
        handleNotifications(req.body);

    }
});

app.post('/create-subscription', async (req, res) => {
    try {
        const accessToken = await getAccessToken();
        const subscription = await createSubscription(accessToken);
        res.json(subscription);
    } catch (error) {
        res.status(500).send('Error creating subscription');
    }
});

app.listen(3000, () => {
    console.log('Server is listening on port 3000');
});