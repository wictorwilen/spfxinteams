# SharePoint Framework Web Part in Microsoft Teams demo

![Works on my machine](https://img.shields.io/badge/Works-on%20my%20machine-green.svg)

**Link**: https://askwictor.com/spfxinteams

Demo solution for Microsoft Ignite 2018 that shows how a SharePoint Framework Web Part can be added as a Microsoft Teams tab.

> **NOTE** ALL CODE AND SETUP AND CONFIGURATION ARE SUBJECT TO CHANGE AND ARE SHARED ONLY FOR DEMONSTRATIONAL PURPOSES AT MICROSOFT IGNITE 2018. IT MIGHT NOT EVEN WORK IN YOUR TENANT SINCE IT ALSO HAS DEPENDENCIES ON TENANT PRE-RELEASE FEATURES.


## How to install and configure

Follow these instructions to install and configure the solution

1. Upload the .sppkg (SPFx package) into the SharePoint App Catalog and allow it to be installed in all site collections
2. Go to the new SharePoint Tenant Admin page and approve the API permission request (Group.ReadWrite.All)
3. Create a new Microsoft Teams team
4. Wait a minute until the SharePoint team site is provisioned
5. In the associated SharePoint team site create a list called `HostedAppConfigList`
6. Add a single line of text column to the list called `HostType`
7. Add a multi line of text column to the list called `HostData`
8. Create a new App in the site of the type Calendar
9. Add a few events to the calendar you just create
10. Add a new page to the site
11. Add the `Calendar` web part to the new page
12. Configure it to use the calendar app you created and verify that the Web Part works in SharePoint
13. In your Microsoft Teams team, go to settings and apps
14. Choose to upload a custom app, and select the `spfxinteams-teams.zip` file
15. Add a new Tab to the Team and choose the *SPFx in Teams* app
16. When configuring the Tab choose the calendar you created in the SharePoint site.
17. Wait a few seconds before you save the changes (if you get a message that says you need to configure the web part, remove the tab and try again, and this time wait an additional few seconds)
18. Verify that the Tab shows the events from the calendar

## How to build

To build and package the solutions, you need to run `npm install` in the *spfx* and *teams* folders respectively.

The *spfx*  folder contains the SharePoint Framework solution and you build that as you normally do.

The *teams*  folder contains a Microsoft Teams project and to package the solution you run `gulp zip` which creates the app package in the *package* folder. (Note; the `gulp manifest` task does not work with this demo as it validates the schema).


(C) Copyright Wictor Wilén, 2018