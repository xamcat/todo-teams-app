# Teams ToDo App

## Summary

This repository contains a simple ToDo app for Microsoft Teams built with the latest [TeamFx toolkit](https://docs.microsoft.com/javascript/api/@microsoft/teamsfx/?view=msteams-client-js-latest). This is a Tab app build with React and Northstar controls. Main features:

1. ToDo items list with ability to add, remove, edit items
1. Local storage for persistance
1. Graph API user search and custom queries
1. Azure function calls

## Demo

![Demo](todo-demo.gif)

## Run the sample

- Install [TeamsFx toolkit](https://docs.microsoft.com/microsoftteams/platform/get-started/prerequisites?tabs=vscode)
- Clone the source code
- Open the cloned folder in Visual Studio Code
- Switch to the teams tab and authenticate yourself at Azure and M365 account
- Select Run > Start Debugging
- In the opened browser, select "Add"

The VS Code will install all the required packages, run the backend and the frontend and starts the browser with Teams running your Tab app.

## Figma early design

Build to reflect the Figma design which can be found [here](https://www.figma.com/file/UHHwGcx8YG30LO6fp5gY21/TODO-App) and quick prototype is below

![Figma Prototype](todo-figma-prototype.png)
