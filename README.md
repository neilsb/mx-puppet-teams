[![Support room on Matrix](https://img.shields.io/matrix/mx-puppet-bridge:sorunome.de.svg?label=%23mx-puppet-bridge%3Asorunome.de&logo=matrix&server_fqdn=sorunome.de)](https://matrix.to/#/#mx-puppet-bridge:sorunome.de)

# mx-puppet-teams
This is a early version of a Microsfot Teams puppeting bridge for matrix. It is based on [mx-puppet-bridge](https://github.com/Sorunome/mx-puppet-bridge) and uses the Microsoft Graph API (Beta)

The bridge does not require teams administrative rights, and is initially aimed at supporting chats, rather than team conversations

## Features

Currently supported and planned features are:

- [X] Double Puppeting
- [X] One to One Chats
- [ ] Group Chats
- [ ] Meeting Chats
- [X] Message Edits (Teams -> Matrix only)
- [ ] Message Deletes
- [ ] Reactions
- [ ] Images
- [ ] Attachments
- [ ] Teams Chats (maybe later....)

## Limitations
As well as the open bugs, there are some limitations in the Microsoft Graph API

 - Chats can only be subscribed to for notifications once the chat has been started.  As a result, the bridge needs to poll for new chats.  This means that while messages in existing chats will be delivered almost instantly, new chats will only appear when polled for  (see config option `teams:newChatPollingPeriod`)
 - Attachments cannot currently be sent to MS Teams via Graph API.
 - Events (e.g. read receipts, presence, typing notificaions, etc) are not currently available via the Graph API
 - Graph API does not updating of existing message text  (Prevents sending message edits from Matrix to Teams)

## Requirements
The Bridge will to be accessible to allow Microsoft Graph API to call webhooks via http(s), and for user oauth authenticaion.  It is strongly recommended that a reverse proxy is used to ensure the endpoints are exposed via https. 

The `oauth:serverBaseUri` setting in config.yaml should be set to the base URI  (e.g.  https://my.domain.com/ or https://me.home.net:2700/, etc)


## Install Instructions (from Source)

*   Set up an Azure Application for authentication (See below for detail)
*   Clone and install:
    ```
    git clone https://github.com/neilsb/mx-puppet-teams.git
    cd mx-puppet-teams
    npm install
    ```
*   Edit the configuration file and generate the registration file:
    ```
    cp sample.config.yaml config.yaml
    ```
    Edit config.yaml and fill out info about your homeserver and Azure Application  
    ```
    npm run start -- -r # generate registration file
    ```
*   Copy the registration file to your synapse config directory.
*   Add the registration file to the list under `app_service_config_files:` in your synapse config.
*   Restart synapse.
*   Start the bridge:
    ```
    npm run start
    ```
*   Start a direct chat with the bot user (`@_msteamspuppet_bot:domain.tld` unless you changed the config).
    (Give it some time after the invite, it'll join after a minute maybe.)
*   Authenticate to your Azure Application, then tell the bot to link you account:
    ```
    link MYTOKEN (see below for details)
    ```
*   Tell the bot user to list the available users: (also see `help`)
    ```
    listusers
    ```
    Clicking users in the list will result in you receiving an invite to the bridged chat.  You will automatically be invited to any chat (new or existing) which is updated via Microsoft Teams 

## Setting up an Azure Application
Note: _The Azure account used to set up the application does **not** have to be the tenant for the MS Teams.  You can create a new Azure account, or use a personal account, to set up the application.  There is no cost to setting up the application._

The following steps should be followed to create the Azure application for authentication and access to the Graph API

1. Log into the Azure Portal (https://portal.azure.com/)
2. Select the **App Registrations** resource
3. Select **New registration**
4. Fill out initial details 
   * Give the application a name (e.g. ms-teams-bridge). 
   * For supported account types select "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
   * For Redirect URL, enter the URL to the OAuth call url (Combine `serverBaseUri` and `redirectPath` from the `oauth` section in config.yaml).  e.g.  https://my.domain.com/msteams/oauth
   * Register the app
5. On the application detail screen take a note of the following value which requires to be entered into the config.yaml file
   * `Application (client) ID`  - Set the `oauth:clientId` to this value
6. Click on the **Add a certifcate or secret**  and add a new client secret.  Set the `oauth:clientSecret` in config.yaml to the secret value
7. Click on API permissions and add the following permissions from the Microsoft Graph delegated permission section
   * `Chat.ReadWrite`
   * `ChatMessage.Read`
   * `ChatMessage.Send`
   * `User.Read`
   * `offline_access`

## Linking your Microsoft account
With the bridge running, visit `serverBaseUri`/login (e.g. https://my.domain.com/login).  This will take you to microsoft login page.  After logging in a 6 digit authoisation code will be displayed. This code should be used to link your matrix account to your microsoft account.

Microsoft require you to revalidate the application every (approx) 90 days. To do this visit the login link above to retrieve a new 6 digit code, then talk to the bot and use the `relink {puppetId} {code}` command.   e.g. `relink 1 abcdef`
