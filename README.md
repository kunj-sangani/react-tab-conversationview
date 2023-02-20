# react-tab-conversationview

## Summary

Have you ever struggled to find or search the conversations from Microsoft Team's channel ?

This sample webpart is developed to display the Microsoft Team's channel's conversation in a simpler way which makes easy to search and filters the new conversations and its replies. Please note that this webpart is designed to use in Microsoft Teams only.

## Features

- SPFx based Team's tab.
- Displays all the New(Parent) Conversations on the top
- Ablity to view all the replies of particular conversation
- Option to go to message or reply
- Find messages based on diffrent filters
    - Body search(free text)
    - Based on sender(from)
    - Based on mentions(who all were mentioned in that message)
    - From and To date
    - With Attachments
- Display options - Chat format vs  Tablular view
- Ability to use same filters on all the replies

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.16.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Following Microsoft Graph permissions needs to be approved after uploading the package in the App Catalog

| Permissions         |
|---------------------|
| ChannelMessage.Read.All           |

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| react-tab-conversationview | [Siddharth Vaghasia](https://github.com/siddharth-vaghasia)  |
| react-tab-conversationview | [Kunj Sangani](https://github.com/kunj-sangani)  |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | Feb 21, 2023 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> As this SPFx webpart only works with in Teams's context, please follow below links to deploy it to tenant and make it available in Microsoft Teams

[Package and Deploy](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-web-part-as-ms-teams-tab#package-and-deploy-your-web-part-to-sharepoint)
[Making the web part availble in Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-web-part-as-ms-teams-tab#make-the-web-part-available-in-microsoft-teams)


## Concept Explored

This extension illustrates the following concepts:

- Developing Team's Tab with SPFx
- Usage of Graph Toolkit in SPFx
- Usage of React North Start library SPFx
- Calling Graph API in SPFx
- Concept of using Teams's Aware Logic in SPFx

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Build Microsoft Teams tab using SharePoint Framework - Tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-web-part-as-ms-teams-tab)
