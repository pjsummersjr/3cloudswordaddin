# 3Clouds Microsoft Word Add-In
## Overview
I created the first version of this Office Add-In for a presentation at SharePoint Saturday New England on October 20, 2018. This version was refactored to leverage the React framework in part so that I could more fully leverage the Office UI Fabric React components but also to make the UI simpler to develop. My original version was initializing DOM components using document.createElement and individually setting CSS, child components and event listeners in the raw JS code. Great learning excercise, but not efficient for make the add-in robust and interesting. 

Although this add-in does not serve any real purpose, I am iterating on it both to continue to leverage it for future presentations but also to continue to enhance my development skills. It may not be perfect, but it's mine ;-)

## To compile and run locally
```
npm run webpack

npm start
```

## Components
### Webpack
Webpack and NPM are used to build this project. In dev mode, the webpack-dev-server is used for serving up the content. One of the key TO-DO's is figuring out how to properly build this so it can easily be hosted in a website for real distribution and usage.

### Typescript
This add-in is written in Typescript

### React
This add-in uses the React framework for the display logic and display state management. I found the Office Yoeman generator too cumbersome and complicated so I opted to build this out from scratch. It builds upon some of the core Webpack + Typescript + React scaffolding I put togehter in another repo and simply incorporates the important add-in elements, like authentication and initalization of the Office.js libraries.

### Office-JS-Helpers
The Office-JS-Helpers library is super awesome library developed by Microsoft to make it easier to authenticate to Azure Active Directory, as well as a few other common auth providers, from within an Office Add-In as well as Microsoft Teams. Actually, this library does other stuff to simplify add-in and Teams extensibility but I used it for auth and it saved my life. I highly recommend checking it out.

(https://github.com/OfficeDev/office-js-helpers)

## Code
The code is organized in typical React-based SPA fashion. I have an index.html (src/index.html) that contains the primary content DOM element (id="root"). However, unlike many client-side applications that are bundled with Webpack, there are a few dependencies placed directly in that file. I have the Office-UI-Fabric CSS files and the Office.js reference:
``` HTML
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js"></script>
    <script src="https://unpkg.com/core-js/client/core.min.js"></script>
    <!-- Office Fabric UI stylles 
        - when using Office Fabric React in an Office Add-In, it appears I need to include these as CSS links
          in the HTML file
    -->
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.component.css">
```
Also worth noting that I set the class of the body using an Office UI Fabric style so that the default fonts flow through the add-in.
``` HTML
<body class="ms-Fabric"><!-- Need to include this body style to set some of the baseline fonts -->
```

The real entry point is index.tsx, but this has some Office add-in specifics that are worth calling out.

## Reference
(https://medium.freecodecamp.org/how-to-use-reactjs-with-webpack-4-babel-7-and-material-design-ff754586f618)
(https://webpack.js.org/guides/typescript/)