Program uses Dymo SDK to connect to, and send xml in order to print qr codes generated from an excel upload file from the client.

XML Label format: https://docs.google.com/document/d/1Hb_1qDJmnaWM7-AfKr6LguLxh2nAoWWo_zR2ZmSq4vA/edit?pli=1&tab=t.0#heading=h.1en9qkdfndsz referenced for xml label objects and configuring label to print with the desired layout

Dymo SDK References: https://github.com/dymosoftware/dymo-connect-framework/blob/master/README.md

I've included the dymo.connect.framework.js if needed. Currently SDK is included through a CDN, likely not the most efficient way to do this. To get the program working, I will leave this method of accessing the SDK. Future updates, the SDK should be installed on the server.

The following steps were used to generate this project:
- Create project file (`NodeConsoleApp2.esproj`).
- Create `launch.json` to enable debugging.
- Install npm packages: `npm init && npm i --save-dev eslint`.
- Create `app.js`.
- Update `package.json` entry point.
- Create `eslint.config.js` to enable linting.
- Add project to solution.
- Write this file.