## directline-bot-extension

A starting point SPFx extension to add a DirectLine Bot Framework bot to your SharePoint pages.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

##Files generated during build
* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources

##Files to be added to SharePoint
* temp/deploy/* - all resources which should be uploaded to a CDN.
* sharepoint/solution/directline-bot-extension.sppkg - the SharePoint app file to be uploaded to your App Catalog

### Test/Build options

##Starts the local server to host the extension.
gulp serve --nobrowser 

##Paste the following at the end of any SharePoint modern page in your tenant to test
## ** Be sure to update the three properties noted in the JSON (DirectLineSecret, BotId, BotName)
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"21d2dffd-4f4e-461c-99d4-047c10b21d19":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"DirectLineSecret":"YOUR DIRECTLINE SECRET GOES HERE", "BotId": "YOUR BOT FRAMEWORK ID GOES HERE", "BotName": "YOUR BOT NAME GOES HERE"}}}

##Make sure you've set up your CDN in your tenant before deplying (https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/hosting-extension-from-office365-cdn)
gulp bundle --ship
gulp package-solution --ship
