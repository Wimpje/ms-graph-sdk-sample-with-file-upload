/* 
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
 *  See LICENSE in the source repository root for complete license information. 
 */

(function () {
  angular
    .module('app')
    .controller('MainController', MainController);

  function MainController($scope, $http, $log, GraphHelper) {
    let vm = this;

    // View model properties
    vm.displayName;
    vm.emailAddress;
    vm.fileName;
    vm.requestSuccess;
    vm.requestFinished;

    // View model methods
    vm.uploadFile = uploadFile;

    vm.login = login;
    vm.logout = logout;
    vm.isAuthenticated = isAuthenticated;
    vm.initAuth = initAuth;

    /////////////////////////////////////////
    // End of exposed properties and methods.

    function initAuth() {
      // Check initial connection status.
      if (localStorage.token) {
        processAuth();
      }
      if (!localStorage.token && localStorage.user) {
        // something's off, make sure user is properly logged out
        GraphHelper.logout();
      }
    }

    // Set the default headers and user properties.
    function processAuth() {

      // let the authProvider access the access token
      authToken = localStorage.token;

      if (localStorage.getItem('user') === null) {

        // Get the profile of the current user.
        GraphHelper.me().then(function (user) {

          // Save the user to localStorage.
          localStorage.setItem('user', angular.toJson(user));

          vm.displayName = user.displayName;
          vm.emailAddress = user.mail || user.userPrincipalName;
        });
      } else {
        let user = angular.fromJson(localStorage.user);

        vm.displayName = user.displayName;
        vm.emailAddress = user.mail || user.userPrincipalName;
      }

    }

    vm.initAuth();

    function isAuthenticated() {
      return localStorage.getItem('user') !== null;
    }

    function login() {
      GraphHelper.login();
    }

    function logout() {
      GraphHelper.logout();
    }

    function uploadFile() {
      var file = document.getElementById('file').files[0]
      GraphHelper.uploadFile(file);
    }

  };
})();