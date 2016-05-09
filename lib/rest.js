/**
 * Interface to server REST API
 *
 * Exposes methods for making requests to the REST API
 *
 */
'use strict';

// Utility for manipulating URL querystrings
//var querystring = require('querystring');

// Node HTTP request library; don't error out on invalid HTTPS cert
//var request = require('request').defaults({rejectUnauthorized: false });
//var request = require('request');

// Our application configuration file
//var config = require( '../credentials.json');

// Stores the token we use to access the server API
//var apiToken = '';
//var requestHeaders = {};

var Client = require('node-rest-client').Client;
 
// Promise library
var Promise = require('bluebird');


function RestClient( url, username, password ) {

  // configure basic http auth for every request 
  var options = { 
    user: username, 
    password: password
  };
 
  this.client = new Client( options );

  this.url = url;

}

RestClient.prototype.get = function( resource ) {
  var me = this;

  return new Promise(function(resolve, reject){

    //console.log( 'Getting ' + me.url + resource );

    me.client.get( me.url + resource + '?count=200&offset=0', function( data, response ) {

      if( response.statusCode == 200 ) {
        resolve( data );
      }
      else if ( response.statusCode == 404 ) {
        resolve( null );
      }
      else {
        reject( response.statusCode );
      }
    });
  });
}

/*

RestClient.prototype.batchGet = function( resource ) {
  var me = this;

  return new Promise(function(resolve, reject){

  var request = {
    "operations": [
    {
        'method': 'GET',
        'path': resource,
        'operation_id': 'interests',
        'params': {} 
    },
    ]
  };
  });
}

*/



RestClient.prototype.post = function( resource, data ) {
  var me = this;

  return new Promise(function(resolve, reject){

    //console.log( 'Posting ' + me.url + resource, data );

    // set content-type header and data as json in args parameter 
    var args = {
      data: data,
      headers: { "Content-Type": "application/json" }
    };

    me.client.post( me.url + resource, args, function( data, response ) {

      if( response.statusCode == 200 ) {
        resolve( data );
      }
      else if ( response.statusCode == 404 ) {
        resolve( null );
      }
      else {
        console.log( response );
        console.log( data );
        reject( response.statusCode );
      }
    });
  });



}

RestClient.prototype.patch = function( resource, data ) {
  var me = this;

  return new Promise(function(resolve, reject){

    //console.log( 'Patching ' + me.url + resource, data );

    // set content-type header and data as json in args parameter 
    var args = {
      data: data,
      headers: { "Content-Type": "application/json" }
    };

    me.client.patch( me.url + resource, args, function( data, response ) {

      if( response.statusCode == 200 ) {
        resolve( data );
      }
      else if ( response.statusCode == 404 ) {
        resolve( null );
      }
      else {
        console.log( response );
        console.log( data );
        reject( response.statusCode );
      }
    });
  });



}

module.exports = RestClient;

