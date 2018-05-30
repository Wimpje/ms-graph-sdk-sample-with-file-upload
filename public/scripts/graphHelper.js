/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/


// About MSGRAPH: https://github.com/microsoftgraph/msgraph-sdk-javascript#usage
// About MSAL: (authentication library): https://docs.microsoft.com/en-us/azure/active-directory/develop/guidedsetups/active-directory-javascriptspa

"use strict";

function createApplication(applicationConfig) {

    var clientApplication = new Msal.UserAgentApplication(applicationConfig.clientID, null, function (errorDesc, token, error, tokenType) {
        // Called after loginRedirect or acquireTokenPopup
    });

    return clientApplication;
}

var clientApplication;

(function () {
  angular
    .module('app')
    .service('GraphHelper', ['$http', function ($http) {

      // Initialize the auth request.
      clientApplication = createApplication(APPLICATION_CONFIG);

      return {

        // Sign in and sign out the user.
        login: function login() {
            clientApplication.loginPopup(APPLICATION_CONFIG.graphScopes).then(function (idToken) {
                clientApplication.acquireTokenSilent(APPLICATION_CONFIG.graphScopes).then(function (accessToken) {
                    localStorage.token = accessToken;
                    window.location.reload();
                }, function (error) {
                    clientApplication.acquireTokenPopup(APPLICATION_CONFIG.graphScopes).then(function (accessToken) {
                        localStorage.token = accessToken;
                    }, function (error) {
                        window.alert("Error acquiring the popup:\n" + error);
                    });
                })
            }, function (error) {
                window.alert("Error during login:\n" + error);
            });
        },
        logout: function logout() {
            clientApplication.logout();
            delete localStorage.token;
            delete localStorage.user;
        },
  

        // Get the profile of the current user.
        me: function me() {
          return graphClient.api('/me').get();
        },

        uploadFile: function uploadFile(file) {
            var reader = new FileReader();
            reader.addEventListener("load", function () {
              return graphClient.api('/me/drive/special/approot/children/'+ file.name +'/content')
                  .put(file, (err, res) => {
                    if (err) {
                      console.log(err, res, file);
                      return;
                    }
                    console.log("We've uploaded your file!");
                  });
            }, false);

            if (file) {
              reader.readAsDataURL(file);
            }
        },

        // Send an email on behalf of the current user.
        sendMail: function sendMail(email) {
          return graphClient.api('/me/sendMail').post({ 'message' : email, 'saveToSentItems': true });
        }
      }
    }]);
})();
