# MojitoLib
**Google Apps Script library for the [Mojito spreadsheet](http://b3devs.blogspot.com/p/about-mojito.html)**

To modify this library to your liking, you must do the following:

1. Clone the [MojitoLib git repo](https://github.com/b3devs/MojitoLib)
1. Install [Node.js](https://nodejs.org/)
1. Install [node-google-apps-script](https://www.npmjs.com/package/node-google-apps-script) (TODO: Need to change this to use Google's [clasp](https://www.npmjs.com/package/@google/clasp) tool instead.)
1. Follow the [quickstart steps](https://www.npmjs.com/package/node-google-apps-script#quickstart) for gapps to get MojitoLib uploaded to your Google Drive.
1. Edit one or more files using you favorite Javascript editor.
1. Build the changes:  ```npm run build```
1. Upload the changes to your Google Drive project: ```gapps push```
1. Open your updated MojitoLib in Google Drive and create a new version (File > Manage Versions menu).
1. Open your copy of Mojito, go to the Script Editor, and [change the MojitoLib library](https://developers.google.com/apps-script/guides/libraries) to the one in *your* Google Drive.
1. When you are happy with your changes, commit the files to your local git repo (git add -u, git commit)
