/* Install dependency before start.
 * Execute 'npm install' if node_modules folder doesn't exist.
 * Execute 'dotnet build -o bin' if MODS binding .dll doesn't exist.
 */
var fs = require('fs');
var path = require('path');
var child = require('child_process');

var argv = process.argv.slice(2);
if (!argv || !argv.length) {
  console.error('Run command not specified, exiting.');
  process.exit(1);
}

var allTips = {
  checkSSHKey: 'Check that your git/SSH key is authorized to connect to private MODS repositories found in package.json.',
  checkDotNet: 'Your .NET Core version must be at least v3.1.',
  runOneClickInstaller: 'Try running the MODS installer command one more time.'
};

function showTips(tips) {
  console.error('Maybe you can fix it with following tip(s).');
  tips.forEach(tip => console.error(`* ${tip}`));
}

var steps = [
  {
    validationPath: '../node_modules',
    command: 'npm install',
    tips: [allTips.checkSSHKey]
  },
  {
    validationPath: '../bin/Microsoft.Azure.WebJobs.Extensions.MODS.dll',
    command: 'dotnet build -o bin',
    tips: [allTips.checkDotNet, allTips.runOneClickInstaller]
  }
];

// Use this function because Node v10 doesn't support recursive option on rmdirSync().
function cleanupPath(target) {
  if (!fs.existsSync(target)) {
    return;
  }

  if (!fs.statSync(target).isDirectory()) {
    fs.unlinkSync(target);
    return;
  }

  var isWindows = process.platform == 'win32';
  if (isWindows) {
    child.execSync(`rmdir /s /q "${target}"`);
  } else {
    child.execSync(`rm -rf "${target}"`);
  }
}

steps.forEach(step => {
  var validationPath = path.resolve(__dirname, step.validationPath);
  if (!fs.existsSync(validationPath)) {
    try {
      child.execSync(step.command, { stdio: 'inherit' });
    } catch {
      console.error('');
      console.error(`Failed to execute '${step.command}'.`);
      showTips(step.tips);
      console.error('');

      /* Remove target path on failure,
       * otherwise it will skip the next run.
       */
      cleanupPath(validationPath);
      process.exit(1);
    }
  }
});

child.spawn(argv[0], argv.slice(1), { stdio: 'inherit', shell: true });
