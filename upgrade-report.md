# Upgrade project react-command-share-to-teams to v1.14.0

Date: 8/12/2022

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.14.0. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.14.0
```

File: [./package.json:14:9](./package.json)

### FN001011 @microsoft/sp-dialog | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-dialog

Execute the following command:

```sh
npm i -SE @microsoft/sp-dialog@1.14.0
```

File: [./package.json:15:9](./package.json)

### FN001013 @microsoft/decorators | Required

Upgrade SharePoint Framework dependency package @microsoft/decorators

Execute the following command:

```sh
npm i -SE @microsoft/decorators@1.14.0
```

File: [./package.json:12:9](./package.json)

### FN001014 @microsoft/sp-listview-extensibility | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-listview-extensibility

Execute the following command:

```sh
npm i -SE @microsoft/sp-listview-extensibility@1.14.0
```

File: [./package.json:16:9](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.14.0
```

File: [./package.json:23:9](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.14.0
```

File: [./package.json:24:9](./package.json)

### FN002009 @microsoft/sp-tslint-rules | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-tslint-rules

Execute the following command:

```sh
npm i -DE @microsoft/sp-tslint-rules@1.14.0
```

File: [./package.json:25:9](./package.json)

### FN006005 package-solution.json metadata | Required

In package-solution.json add metadata section

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "react-command-share-to-teams description"
      },
      "longDescription": {
        "default": "react-command-share-to-teams description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.14.0"
  }
}
```

File: [./.yo-rc.json:5:5](./.yo-rc.json)

### FN014008 Hosted workbench type in .vscode/launch.json | Recommended

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration

```json
{
  "configurations": [
    {
      "type": "pwa-chrome"
    }
  ]
}
```

File: [.vscode\launch.json:28:7](.vscode\launch.json)

### FN001008 react | Required

Upgrade SharePoint Framework dependency package react

Execute the following command:

```sh
npm i -SE react@16.13.1
```

File: [./package.json:19:9](./package.json)

### FN001009 react-dom | Required

Upgrade SharePoint Framework dependency package react-dom

Execute the following command:

```sh
npm i -SE react-dom@16.13.1
```

File: [./package.json:20:9](./package.json)

### FN001022 office-ui-fabric-react | Required

Install SharePoint Framework dependency package office-ui-fabric-react

Execute the following command:

```sh
npm i -SE office-ui-fabric-react@7.174.1
```

File: [./package.json:11:5](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Remove SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm un -D @microsoft/sp-webpart-workbench
```

File: [./package.json:26:9](./package.json)

### FN002015 @types/react | Required

Upgrade SharePoint Framework dev dependency package @types/react

Execute the following command:

```sh
npm i -DE @types/react@16.9.51
```

File: [./package.json:27:9](./package.json)

### FN002018 @microsoft/rush-stack-compiler-3.9 | Required

Install SharePoint Framework dev dependency package @microsoft/rush-stack-compiler-3.9

Execute the following command:

```sh
npm i -DE @microsoft/rush-stack-compiler-3.9@0.4.47
```

File: [./package.json:22:5](./package.json)

### FN007002 serve.json initialPage | Required

Update serve.json initialPage URL

```json
{
  "initialPage": "https://enter-your-SharePoint-site/_layouts/workbench.aspx"
}
```

File: [./config/serve.json:1:1](./config/serve.json)

### FN012017 tsconfig.json extends property | Required

Update tsconfig.json extends property

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-3.9/includes/tsconfig-web.json"
}
```

File: [./tsconfig.json:2:3](./tsconfig.json)

### FN014007 Local workbench in .vscode/launch.json | Recommended

In the .vscode/launch.json file, remove the local workbench launch configuration

```json

```

File: [.vscode\launch.json:11:7](.vscode\launch.json)

### FN015007 config\copy-assets.json | Required

Remove file config\copy-assets.json

Execute the following command:

```sh
rm "config\copy-assets.json"
```

File: [config\copy-assets.json](config\copy-assets.json)

### FN024001 Create .npmignore | Required

Create the .npmignore file


File: [./.npmignore](./.npmignore)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm un -D @microsoft/sp-webpart-workbench
npm i -SE @microsoft/sp-core-library@1.14.0 @microsoft/sp-dialog@1.14.0 @microsoft/decorators@1.14.0 @microsoft/sp-listview-extensibility@1.14.0 react@16.13.1 react-dom@16.13.1 office-ui-fabric-react@7.174.1
npm i -DE @microsoft/sp-build-web@1.14.0 @microsoft/sp-module-interfaces@1.14.0 @microsoft/sp-tslint-rules@1.14.0 @types/react@16.9.51 @microsoft/rush-stack-compiler-3.9@0.4.47
npm dedupe
rm "config\copy-assets.json"
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

In package-solution.json add metadata section:

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "react-command-share-to-teams description"
      },
      "longDescription": {
        "default": "react-command-share-to-teams description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.14.0"
  }
}
```

#### [.vscode\launch.json](.vscode\launch.json)

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration:

```json
{
  "configurations": [
    {
      "type": "pwa-chrome"
    }
  ]
}
```

In the .vscode/launch.json file, remove the local workbench launch configuration:

```json

```

#### [./config/serve.json](./config/serve.json)

Update serve.json initialPage URL:

```json
{
  "initialPage": "https://enter-your-SharePoint-site/_layouts/workbench.aspx"
}
```

#### [./tsconfig.json](./tsconfig.json)

Update tsconfig.json extends property:

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-3.9/includes/tsconfig-web.json"
}
```

#### [./.npmignore](./.npmignore)

Create the .npmignore file:

```text
!dist
config

gulpfile.js

release
src
temp

tsconfig.json
tslint.json

*.log

.yo-rc.json
.vscode

```
