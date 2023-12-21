/* global require, process, console */

const fs = require("fs");
const host = "outlook";
const hosts = ["excel", "outlook", "powerpoint", "word"];
const path = require("path");
const util = require("util");
const testPackages = [
  "@types/mocha",
  "@types/node",
  "mocha",
  "office-addin-mock",
  "office-addin-test-helpers",
  "office-addin-test-server",
  "ts-node",
];
const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);

async function modifyProjectForSingleHost(host) {
  if (!host) {
    throw new Error("The host was not provided.");
  }
  if (!hosts.includes(host)) {
    throw new Error(`'${host}' is not a supported host.`);
  }
  await convertProjectToSingleHost(host);
  await updatePackageJsonForSingleHost(host);
  await updateLaunchJsonFile();
}

async function convertProjectToSingleHost(host) {
  // Delete the test folder
  deleteFolder(path.resolve(`./test`));

  // Delete the .github folder
  deleteFolder(path.resolve(`./.github`));

  // Delete CI/CD pipeline files
  deleteFolder(path.resolve(`./.azure-devops`));

  // Delete repo support files
  await deleteSupportFiles();
}

async function updatePackageJsonForSingleHost(host) {
  // Update package.json to reflect selected host
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Remove 'engines' section
  delete content.engines;

  // Remove scripts that are unrelated to the selected host
  Object.keys(content.scripts).forEach(function (key) {
    if (
      key === "convert-to-single-host" ||
      key === "start:desktop:outlook"
    ) {
      delete content.scripts[key];
    }
  });

  // Remove test-related scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("test")) {
      delete content.scripts[key];
    }
  });

  // Remove test-related packages
  Object.keys(content.devDependencies).forEach(function (key) {
    if (testPackages.includes(key)) {
      delete content.devDependencies[key];
    }
  });

  // Write updated json to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updateLaunchJsonFile() {
  // Remove 'Debug Tests' configuration from launch.json
  const launchJson = `.vscode/launch.json`;
  const launchJsonContent = await readFileAsync(launchJson, "utf8");
  const regex = /(.+{\r?\n.*"name": "Debug (?:UI|Unit) Tests",\r?\n(?:.*\r?\n)*?.*},.*\r?\n)/gm;
  const updatedContent = launchJsonContent.replace(regex, "");
  await writeFileAsync(launchJson, updatedContent);
}

function deleteFolder(folder) {
  try {
    if (fs.existsSync(folder)) {
      fs.readdirSync(folder).forEach(function (file) {
        const curPath = `${folder}/${file}`;

        if (fs.lstatSync(curPath).isDirectory()) {
          deleteFolder(curPath);
        } else {
          fs.unlinkSync(curPath);
        }
      });
      fs.rmdirSync(folder);
    }
  } catch (err) {
    throw new Error(`Unable to delete folder "${folder}".\n${err}`);
  }
}

async function deleteSupportFiles() {
  await unlinkFileAsync("CONTRIBUTING.md");
  await unlinkFileAsync("LICENSE");
  await unlinkFileAsync("README.md");
  await unlinkFileAsync("SECURITY.md");
  await unlinkFileAsync("./convertToSingleHost.js");
  await unlinkFileAsync(".npmrc");
  await unlinkFileAsync("package-lock.json");
}

/**
 * Modify the project so that it only supports a single host.
 * @param host The host to support.
 */
modifyProjectForSingleHost(host).catch((err) => {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
  process.exitCode = 1;
});