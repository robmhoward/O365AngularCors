var spTenantDomain = "junction.sharepoint.com";
var spBaseUrl = "https://" + spTenantDomain + "/portals/hub/_api/videoService/";
var odbBaseUrl = "https://junction-my.sharepoint.com/"
var exchangeBaseUrl = "https://outlook.office365.com/";
var azureAdBaseUrl = "https://graph.windows.net/";

var o365CorsApp = angular.module("o365CorsApp", ['ngRoute', 'AdalAngular'])
	.factory('o365CorsFactory', ['$http', function ($http) {
		var factory = {};

		$http.defaults.useXDomain = true;
		
		factory.getCalendarEvents = function() {
			return $http.get(exchangeBaseUrl + 'api/v1.0/me/events?$top=10');
		}

		factory.getMessages = function() {
			return $http.get(exchangeBaseUrl + 'api/v1.0/me/messages?$top=10');
		}

		factory.getFiles = function() {
			return $http.get(odbBaseUrl + '_api/v1.0/me/files?$top=10')
		}

		return factory;
	}]);

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
			endpoints: {}
			//cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost. 
	};
	adalConfig.endpoints["https://" + spTenantDomain + "/portals/hub/_api/"] = "https://" + spTenantDomain;
	adalConfig.endpoints[odbBaseUrl] = odbBaseUrl;
	adalConfig.endpoints[exchangeBaseUrl] = exchangeBaseUrl;
	adalConfig.endpoints[azureAdBaseUrl] = azureAdBaseUrl;

	adalProvider.init(adalConfig, $httpProvider); 

}]);


o365CorsApp.controller("HomeController", function($scope, $q, o365CorsFactory) {

	$scope.files = [{name: "Loading..."}];
	$scope.messages = [{Subject: "Loading..."}];
	$scope.calendarEvents = [{Subject: "Loading..."}];

	o365CorsFactory.getFiles().then(function(response) {
		$scope.files = response.data.value;
	});

	o365CorsFactory.getCalendarEvents().then(function(response) {
		$scope.calendarEvents = response.data.value;
	});

	o365CorsFactory.getMessages().then(function(response) {
		$scope.messages = response.data.value;
	});


	/*
	var requests = {
		calendarEvents: o365CorsFactory.getCalendarEvents(),
		messages: o365CorsFactory.getMessages(),
		files: o365CorsFactory.getFiles()
	};

	$q.one(requests).then(function (responses) {
		console.log("Requests completed");
		$scope.calendarEvents = responses.calendarEvents.data.value;
		$scope.messages = responses.messages.data.value;
		$scope.files = responses.files.data.value;
		console.log($scope.files.length + " files returned");
	});
*/
});