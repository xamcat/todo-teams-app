const fs = require("fs");
const path = require("path");
const { argv } = require("process");

let launchTaskName;
const validateType = argv[2];

const PRODUCTION_LAUNCH_TASK_NAME = "Debug (Teams Web Client)";
const LOCAL_DEBUG_LAUNCH_TASK_NAME = "Local Debug Tab App (Teams Web Client)";
const CONSOLE_COLOR_RED = "\x1b[31m";
const CONSOLE_COLOR_YELLOW = "\x1b[33m";
const CONSOLE_COLOR_BLUE = "\x1b[34m";

if (validateType === "azure") {
    launchTaskName = PRODUCTION_LAUNCH_TASK_NAME;
}
else if (validateType === "local") {
    launchTaskName = LOCAL_DEBUG_LAUNCH_TASK_NAME
}

if (launchTaskName) {
    const launchJsonString = fs.readFileSync(path.join(__dirname, "launch.json"), "utf8");
    const launchJson = JSON.parse(launchJsonString);
    const config = launchJson.configurations.find(config => config.name === launchTaskName);
    if (config) {
        if (config.url.match(/\{.*?[Tt]eamsAppId\}/)) {
            if (launchTaskName == LOCAL_DEBUG_LAUNCH_TASK_NAME) {
                console.log(CONSOLE_COLOR_RED, "launch.json:1:1: error: You need to setup local development environment first before launching Teams App\n");
                console.log(CONSOLE_COLOR_YELLOW, "In the command palette, run 'MODS - Create environment -> Local' to create local development environment\n");
            }
            else if (launchTaskName == PRODUCTION_LAUNCH_TASK_NAME) {
                console.log(CONSOLE_COLOR_RED, "launch.json:1:1: error: You need to provision and deploy Azure resources first before launching Teams App\n");
                console.log(CONSOLE_COLOR_YELLOW, "In the command palette, run 'MODS - Create environment -> Azure' and 'MODS - Deploy All (frontend and backend)' to provision and deploy Azure resources\n");
            }

            console.log(CONSOLE_COLOR_YELLOW, "Step by step instructions can be found here:", CONSOLE_COLOR_BLUE, "https://aka.ms/MODSPrivatePreview/guide\n");
            process.exit(1);
        }
    }
}
