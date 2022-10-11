# EWS Middleware

[TOC]

## Prerequisites

- Node.js 12 ( Node.js 10 would also work, but 12 is used when develop & test)
- ngrok (Used to start HTTPS proxy server for autodiscover)
- Postman for verification



## Configuration

Configuration can be done via `config.js` and via environment variables.

| Name                | Environment Variable   | Description                                                  | Example                                        |
| ------------------- | ---------------------- | ------------------------------------------------------------ | ---------------------------------------------- |
| port                | PORT                   | The port of app to run                                       | 3000                                           |
| host                | HOST                   | The host of app to run                                       | localhost                                      |
| serverUrl           | SERVER_URL             | The server url. Used to construct external EWS url of auto discover response. | [http://localhost:3000](http://localhost:3000) |
| maxRequestBodySize  | MAX_REQUEST_BODY_SIZE  | The max request payload size                                 | 5MB                                            |
| gracePeriodForClose | GRACE_PERIOD_FOR_CLOSE | Grace period for app close, in milliseconds.                 | 5000                                           |
| keepAliveInterval   | KEEP_ALIVE_INTERVAL    | Interval time to keep streaming connection alive, in milliseconds | 30000                                          |



## Run EWS Middleware services

```bash
# Install dependencies
npm install

# Lint code
npm run lint

# Build typescript code
npm run build

# Clean built code
npm run clean

# Run middleware services
npm start
```



When you see something like `Server is running at ...` the server is started successfully.



## Validation

The server will serve static wsdl and xsd files at:

- http://localhost:3000/EWS/Services.wsdl
- http://localhost:3000/EWS/messages.xsd
- http://localhost:3000/EWS/types.xsd



Then import `doc/ews.postman_collection.json` and `doc/ews.postman_environment.json` into Postman to verify SOAP services.

The console will log incoming SOAP request. E.g. for `GetClientAccessToken` operation the log will be:

```json
GetClientAccessTokenType {
  TokenRequests: NonEmptyArrayOfClientAccessTokenRequestsType {
    TokenRequest: [
      ClientAccessTokenRequestType {
        Id: '1C50226D-04B5-4AB2-9FCD-42E236B59E4B',
        TokenType: 'CallerIdentity'
      },
      ClientAccessTokenRequestType {
        Id: '1C50226D-04B5-4AB2-9FCD-42E236B59E4B',
        TokenType: 'ExtensionCallback'
      }
    ]
  },
  SOAPHeaders: { RequestServerVersion: { Version: 'Exchange2016' } }
}
```

Here the `GetClientAccessTokenType` is the type of the parsed JS object, it contains a `TokenRequests` field which is type of `NonEmptyArrayOfClientAccessTokenRequestsType`, which in turn contains a `TokenRequest` field which is an array. Then the array contains 2 items, each is type of `ClientAccessTokenRequestType`.

You can see here the class type are properly constructed (which means eligible to use `instanceof ` to check type).



## Verify Outlook

Currently only Mac Outlook is supported.

At first for autodiscover to work, you need use `ngrok` to start a HTTPS proxy server.

```bash
ngrok http 3000
# Then the console will ouput something like:
# Forwarding  https://xxxxxxxx.ngrok.io -> http://localhost:3000
```

The ngrok panel at `http://localhost:4040/inspect/http` will log the autodiscover requests.



Then in Outlook accounts setting, create two Exchange accounts: `user1@xxxxxxxx.ngrok.io` and `user2@xxxxxxxx.ngrok.io` with any password.

Then you can send emails between the above two users. Streaming events subscription is implemented so once an email is sent, the receiver will immediately receive it.

File attachment, email attachment, inline picture, thread conversations are also supported and demostrated.

See video: https://youtu.be/uxP98lJNq3o

 
 
## Setting up SSL on your localhost

The free version of ngrok has a limitation of 20 connections per minute. EWS can hit that limit if you have more than a small amount of test data. After hitting the limit nothing is sent for the remainder of the minute. This can cause sync issues. 

You can setup EWS to use SSL (https) so that you can configure your clients to go directly to EWS server running on your local machine. 

 
### Create a self signed certificate

1. In the Keychain Access app on your Mac, choose Keychain Access > Certificate Assistant > Create a Certificate.
1. Enter a name for the certificate.
1. Set Identity type to `Self Signed Root`
1. Set Certificate Type to `SSL Server`
1. Check `Let me override defaults`. Click Continue until you get the email field and update it. You may also update common name, organization, city, state if you wish. When you get to the panel with Key Size, set it to 4096. Take defaults for all other fields.  
1. You will see a certificate in your login keychain with the common name you choose (default is the name of your machine, e.g. MyName-MacBook-Pro.local). Right click on it and select Export.
1. Set file format to .p12 and save it somewhere on your machine. You will need to enter a passphrase. Be sure the remember this. 
 

### Configure EWS to use SSL

Setup a launch configuration in VS Code. Under Run add or edit a configuration and update your launch configuration to set the following environment variables:
   1. PROTOCOL - set to https
   1. PFX - set the the location of the p12 file containing your local self signed certificate you created above. This is relative to the root of the EWS middleware code. 
   1. PASSPHRASE - the passphrase you used when exporting your self signed certificate.
   1. SERVER_URL - the url for the EWS server. Be sure to use https. 
   1. PORT - the port used by the EWS server. Normally 443 for SSL. 

  This is an example of a launch configuration with the proper environment variables set:
  ```json
  {
      "type": "node",
      "request": "launch",
      "name": "My Launch Configuration",
      "skipFiles": [
          "<node_internals>/**"
      ],
      "program": "${workspaceFolder}/index.js",
      "preLaunchTask": "tsc: build - tsconfig.json",
      "outFiles": [
          "${workspaceFolder}/dist/**/*.js"
      ],
      "env": { "PROTOCOL": "https", "PFX": "../../../localhostCert.p12", "PASSPHRASE": "xxxxx", "SERVER_URL": "https://localhost", "PORT": "443"}
  }
  ```

When you start EWS you should see it start cleanly on the correct url:
  ```
  Server is running at https://[::1]:443
  ```

When you are running with https, you can configure your exchange account with an email address using a domain that maps to 127.0.0.1 in your hosts file. Outlook does require that the domain have a dot so "localhost" can't be used. 

## Testing

You can use Test Explorer in VSCode to run the unit test or you can run them from the command line by running `npm run test:unit`. See [Writing Unit Tests](docs/unittest.md) for more information. 
 
Use AcceptanceTest.postman_environment.json to set up the environment for the Postman Acceptance Tests.

`npm test` will run all the unit tests and the Postman acceptance test. To run a particular folder use `--folder` and pass the folder name as the parameter, like in the example given below.
`npm test -- --folder "User Configuration Test"` which is the equivalent for `newman run doc/ews.postman_collection.json -e doc/ews.postman_environment.json --folder "User Configuration Test"`


When submitting a Pull Request, you can run `npm run testPR` to run the Postman acceptance tests and get an output report you can include in the PR. 

To run the server in different time zone than the time zone of your machine, add setting the enviroment variable TZ to the start npm script in package.json. For example:
```
"start": "env TZ=Europe/Moscow node -r source-map-support/register ."
```
Time zone names can be found here: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones

## About strong-soap

The `strong-soap` library is used to convert the incoming SOAP request to json objects and json object responses back to SOAP responses. strong-soap is patched to resolve following problems:

- Does not support output "xml:lang" attribute, which is used by `UserOofSettings.InternalReply`, see:

  https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/internalreply#attributes

- It always output pretty formatted xml string, this will increase bandwidth usage 

- When parsing xml it does not parse string value to built-in type (like number, boolean, Date) according to xsd schema

- When parsing xml it does not call proper class constructor for custom object type. The class type would be necessary since sometimes we may need class type to distinguish (E.g. an array contains items having different sub class types)

- When parsing xml it sets the attributes values into a nested field (`$attributes`). From the aspect of coding this causes inconvenience (E.g. it will be easier to code `obj.Id` instead of `obj.$attribute.Id` )

- Lack support for [XSD element substitution](https://www.ibm.com/support/knowledgecenter/SSV2LR/com.ibm.wbpm.wid.data.doc/transform/topics/csubgroup.html) via `substitutionGroup` 

- Will include an additional field in model objects that contain Date properties. In certain cases EWS may include date strings with no time zone information which strong-soap will set as a Date object. The Date object will always be converted to the server's time zone when it is constructed if the date string in the SOAP request included time zone information. If it did not contain time zone information, the Date object will be in the server's time zone with the exact date specified in the SOAP request. However the time zone for the date may be in a seperate field in the SOAP request. Therefore the server code which process the request may need to know if the Date object was constructed with a time zone or not.  In able to determine this, the strong-soap code has been patch to include an extra field for Date properties in the SOAP request object which is the string value from the SOAP request date property. The name of this extra field will be the name of the Date property + 'String'. For example in the CalendarItemType, `Start` is defined as a Date, so a SOAP request which contains a CalendarItemType will contain a field `Start` which is a Date object and `StartString` which is a string. 

The patch is located at `patches/strong-soap+1.22.1.patch`, it is **automatically** applied when run `npm install` using `patch-package` lib, you don't need manually apply it.

If you need to update the patch, you can edit the files under node_modules/strong-soap, then run `npx patch-package strong-soap`. See the [patch-package README](https://github.com/ds300/patch-package#readme) for more information
