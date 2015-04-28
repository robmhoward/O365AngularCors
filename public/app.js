var o365CorsApp = angular.module("o365CorsApp", ['ngRoute', 'AdalAngular'])

o365CorsApp.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', function ($routeProvider, $httpProvider, adalProvider) {  
	$routeProvider
		.when('/',
			{
				controller: 'HomeController',
				templateUrl: 'partials/home.html',
				requireADLogin: true
			})
		.otherwise({redirectTo: '/' });

	var adalConfig = {
		tenant: 'common', 
		clientId: '2005547b-b711-477d-8fdd-c25ed8003191', 
		extraQueryParameter: 'nux=1',
		endpoints: {
			"https://junction-my.sharepoint.com/_api/v1.0": "https://junction-my.sharepoint.com/",
			"https://outlook.office365.com/api/v1.0": "https://outlook.office365.com/",
			"https://www.onenote.com/": "https://onenote.com/",
			"https://graph.windows.net/": "https://graph.windows.net/"
		}
	};
	
	adalProvider.init(adalConfig, $httpProvider); 

}]);


o365CorsApp.factory('o365CorsFactory', ['$http', function ($http) {
	var factory = {};

	$http.defaults.useXDomain = true;
	
	factory.getCalendarEvents = function() {
		return $http.get('https://outlook.office365.com/api/v1.0/me/events?$top=10');
	}

	factory.getMessages = function() {
		return $http.get('https://outlook.office365.com/api/v1.0/me/messages?$top=10');
	}

	factory.getFiles = function() {
		return $http.get('https://junction-my.sharepoint.com/_api/v1.0/me/files?$top=10');
	}

	factory.getNotebooks = function() {
		return $http.get('https://www.onenote.com/api/beta/me/notes/notebooks?$top=10');
	}

	factory.getGroups = function() {
		return $http.get('https://graph.windows.net/me/memberOf?$top=10&api-version=1.5');
	}

	return factory;
}]);

o365CorsApp.controller("HomeController", function($scope, $q, o365CorsFactory) {

	$scope.files = [{name: "Loading..."}];
	$scope.messages = [{Subject: "Loading..."}];
	$scope.calendarEvents = [{Subject: "Loading..."}];
	$scope.notebooks = [{name: "Loading..."}];
	$scope.groups = [{displayName: "Loading..."}];

	o365CorsFactory.getFiles().then(function(response) {
		$scope.files = response.data.value;
	});
	
	o365CorsFactory.getCalendarEvents().then(function(response) {
		$scope.calendarEvents = response.data.value;
	});
	
	o365CorsFactory.getMessages().then(function(response) {
		$scope.messages = response.data.value;
	});

	o365CorsFactory.getNotebooks().then(function(response) {
		$scope.notebooks = response.data.value;
	});

	o365CorsFactory.getGroups().then(function(response) {
		$scope.groups = response.data.value;
	});

});