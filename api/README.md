# Build Teams Application Backend with MODS and Azure Functions

When building a Teams application, MODS provides an option for you to add a backend API to develop server-side logics so that you can easily build your systems to react to a series of critical events. The API you added is actually an [Azure Functions](https://docs.microsoft.com/en-us/azure/azure-functions/) project that handles HTTP requests from Tabs, and you can customize it according to your requirements.

## Prerequisites

To start enjoying full functionalities to develop an API with Azure Functions for your Teams Application, you need to:
- Install [Azure Functions Core Tools](https://docs.microsoft.com/en-us/azure/azure-functions/functions-run-local?tabs=windows%2Ccsharp%2Cbash).
- Install [MODS Server SDK Package](https://aka.ms/MODSPrivatePreview/server-sdk).
- Add an API during project creation or using command, see [MODS User Manual](https://mods-landingpage-web.azurewebsites.net/md/guide/index).

## Develop

By default, MODS will provide template code for you to get started. The starter code handles calls from your Teams App client side, initializes the MODS server SDK to access current connected user information and prepares a pre-authenticated Microsoft Graph Client for you to access more user's data. You can modify the template code with your custom logics or add more functions with `HTTPTrigger` by running command `MODS - Add Resource: Azure Function`. Read on [Azure Functions developer guide](https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference) for more development resources.

## Trigger Function

- Invoking MODS Client SDK API `callFunction()` from Tabs.
- Sending an HTTP request to the service. However, MODS binding always checks the SSO token of
  received HTTP request before function handles the request. Thus, requests without a valid SSO token would cause function responses HTTP error 500.

## Debug Locally

You can follow below steps to debug your Azure Function locally:

- Open the root folder of the Teams App project with Visual Studio Code.
- Run `MODS - Create environment` and then select `Local`.
- Open a terminal and change the directory to the `api` folder.
- Execute `npm run start`. It will install dependency and launch a service with Azure Functions Core Tools.
- In the `Run` panel of Visual Studio Code, switch the debug configuration to `Local Debug Function` and press `F5`. Then Visual Studio Code should attach to the worker process.
- You can also launch `Local Debug Tab App` in the same Visual Studio Code window. Then you can debug the two components simultaneously. We suggest you debug this project along with `Local Debug Tab App`, thus you can trigger HTTP request from Tabs with MODS Client SDK.
- After stop debugging, please manually terminate the process in the terminal.

## Deploy to Azure

- Provision Azure environment by running command `MODS - create environment` and choose `Azure`.
- Deploy your project to the Azure Function App by running command - `MODS - Deploy teams app backend with Azure Function`.
- You can also run command `MODS - Deploy All (frontend and backend)` to trigger the deployment along with Tabs.