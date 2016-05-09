/**
 * Utility to load data from an Excel file to a MailChimp List
 *
 */
'use strict';

// Library for communicating with the telematics server
var RestClient = require('./lib/rest');

// Library for reading/writing excel files
var XLSX = require('xlsx');

var Promise = require('bluebird');

var chalk = require('chalk');


// File system interface
//var fs = require('fs');

// File name searching utiility
//var glob = require( 'glob');

// Utility library
var _ = require('underscore');

var crypto = require('crypto');

// store a bunch of global data for easy access
var g = {
  listId: null,
  categories: {}
};

// Retrieve the command line arguments and store them in an object:
// example: -x 3 -y 4 -n5 -abc --beep=boop foo bar baz
// yields: { _: [ 'foo', 'bar', 'baz' ], x: 3, y: 4, n: 5, a: true,
//      b: true, c: true, beep: 'boop' }
var argv = require( 'minimist' )( process.argv.slice( 2 ) );


// Check for the help option and if requested, display help
if( 'h' in argv )
{
    log.info(
        '\nMailChimp List Loader Utility\n\n',
        'Arguments:  [file.xlsx]',

        // Help
        '-h               : Print help\n',
        '\n'
         );

    process.exit( 0 );
}


// Create the object used to access MailChimp
var rest = new RestClient( 
  'https://us12.api.mailchimp.com/3.0/',
  'dontmatter', 
  'enter-your-api-key-here'
  );

// Store the contents of the Excel Sheet
var excelList;



/**
 * Load an excel workbook as a json object
 *
 * @param  {string} file filename
 * @return {object}      JSON array, or null if the workbook could not be read
 */
function readWorkbook( file ) {

    console.log( 'Reading ' + file );

    try {
      var workbook = XLSX.readFile( file );
      var worksheet = workbook.Sheets['IP Adults'];

      return XLSX.utils.sheet_to_json( worksheet );
    }
    catch(e) {
      return null;
    }

}


function validateEmail(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}



/**
 * Adds a unit membership to the subscriber, if it does not already exist
 */
function updateUnit( sub, type, unitNumber ) {

  return new Promise(function(resolve, reject) {

    if( sub ) {
      // get existing unit membership
      var field = '';

      switch( type ) {
        case 'Troop':
          field  = 'TROOPS';
          break;
        case 'Pack':
          field = 'PACKS';
          break;
        case 'Crew':
          field = 'CREWS';
          break;
        case 'Ship':
          field = 'SHIPS';
          break;
        case 'Team':
          field = 'TEAMS';
          break;

        default:
          reject( new Error ('Unknown unit type'));
          return;
          break;

      }

      var units = sub.merge_fields[field].trim();
      var updated = false;

      var unitList = [];
      if( units > '') {
        unitList = units.split(/[\s,]+/);

      }
      
      if( -1 == unitList.indexOf( unitNumber )) {
        unitList.push( unitNumber );
        updated = true;
      }

      if( updated ) {

       var mergeFields = {};

        mergeFields[field] = unitList.join(',');
        sub.merge_fields[field] = mergeFields[field];
 
        console.log( 'update unit ' + sub.email_address + ' ' + type + ' ' + unitNumber );
        
        rest.patch( 'lists/' + g.listId + '/members/' + sub.id, {
          merge_fields: mergeFields,
          }

        )
        .then( function( data ) {
          resolve( sub );
        })
        .catch( function (err) { 
          console.log( chalk.red( 'Failed to update unit ' + sub.email_address ));
          reject(err ); 
        });
      }
      else {
        resolve( sub );
      }      
    }
    else {
      resolve( null );
    }
  });
}

/**
 * Updates the user's position (rank)
 */
function updateRank( sub, therank ) {

  return new Promise(function(resolve, reject) {
    if( sub ) {

      var rank = _.findWhere( g.Positions, { title: therank });

      if( rank ) {

        if( !sub.interests[rank.id]  ) {
          console.log( 'update position ' + sub.email_address + ' ' + rank.id );

          sub.interests[rank.id] = true;

          var interests = {};

          interests[rank.id] = true;

          rest.patch( 'lists/' + g.listId + '/members/' + sub.id, {
            interests: interests,
            }

          )
          .then( function( data ) {
            resolve( sub );
          })
          .catch( function (err) { 
            console.log( chalk.red( 'Failed to update position ' + sub.email_address ));
            reject(err ); 
          });
        }
        else {
          // no need to update
          resolve(sub);
        }

      }
      else {
        reject( new Error( 'Unknown rank ' + therank ));

      }
    }
    else {
      resolve( null );
    }
  });
}


/**
 * Determines the MailChimp member record for the specified email address
 *
 * If not found, creates a new member and provides that.
 */
function findOrCreateSubscriber( email, first, last ) {

  return new Promise(function(resolve, reject) {

    console.log( chalk.blue( first + ' ' + last + ': ' + email ));
    email = email.toLowerCase().trim();
    
    if( email > '' && validateEmail(email) ) {

      //console.log( 'looking up ' + email);

      var hash = crypto
        .createHash('md5')
        .update( email )
        .digest('hex');

      rest.get( 'lists/' + g.listId + '/members/' + hash)
        .then( function( data ) {

          var actions = [];

          if( !data ) {
            // subscriber is not in mailchimp - create
            rest.post( 'lists/' + g.listId + '/members', {
              email_address: email,
              email_type: 'html',
              status: 'subscribed',
              merge_fields: {
                FNAME: first,
                LNAME: last,
                NEWSLETTER: 'No'
              },

            })
            .then( function( data ) {
              resolve( data );
            })
            .catch( function (err) { 
              console.log( chalk.red( 'Failed to add ' + email + ' ' + first + ' ' + last ));
              reject(err ); 
            });
          }
          else {
            resolve( data );
          }

        })
        .catch( function( err ) { reject( err ); });
      }
      else {
        // no email address supplied; can't do much
        console.log( chalk.yellow( 'Skipping ' + first + ' ' + last + ': No email address'));
        resolve( null );
      }
  });
}

/**
 * Inspects the row from the excel sheet and updates MailChimp as necessary
 */
function updateSubscriber( person ) {

  return new Promise(function(resolve, reject) {

    var email = person['Registrant Home E-Mail'];
    var unitType = person['Unit Type'];
    var unit = person['Unit No'];
    var first = person['First Name'];
    var last = person['Last Name'];
    var unitRank = person['Unit Rank'];

    if( !email )
    {
      // silently ignore if no email address
      resolve();
    }
    else {
      findOrCreateSubscriber( email, first, last )
        //.delay( 500)
        .then( function( sub ) { return updateUnit( sub, unitType, unit );})  
        .then( function( sub ) { return updateRank( sub, unitRank );})  

        // Done!
        .then( function( sub ) { resolve(); })
        .catch( function( err ) { reject( err); });
      
    }
  });

}


/**
 * Load the interests in the specified MailChimp Interest Category
 * 
 * store them globally in the g object for later us
 */
function getInterests( interest ) {
  return new Promise(function(resolve, reject){


    var category = _.findWhere( g.categories, { title: interest } );

    if( category ) {

      rest.get( 'lists/' + g.listId + '/interest-categories/' + category.id + '/interests')
        .then( function (data) {

          var list = _.map( data.interests, function( obj ) {
            return { id: obj.id, title: obj.name };
          });

          console.log( 'got ' + list.length + ' ' + interest + 's ' );
          g[interest] = list;

          resolve( list );
        })
        .catch( function( err ) {
          reject( err );
        });
    }
    else {
      reject( new Error( 'Interest not found'));
    }

  });
}

/**
 * Main utility
 *
 * First we have to log into the server 
 * Then we check for stuff that needs to be updated
 *
 */

// Read the excel file
excelList = readWorkbook( argv._[0] );
console.log( 'Excel sheet has ' + excelList.length + ' entries');


rest.get('lists')
  .then( function( data ) {

    var listId = null;

    if( data ) {
      var index = _.findIndex(data.lists, function(voteItem) { 
        return voteItem.name === 'Indian Prairie District Adults' 
      });
  
      if( index > -1 ) {
        g.listId = data.lists[index].id;
        
      }
      else {
        throw new Error( 'Target list not found');
      }
    }
    else {
      throw new Error( 'Unable to retrieve lists');
    }

    //return { listId: listId };
  })
  .then( function( info ) { return rest.get( 'lists/' + g.listId + '/interest-categories'); })

  // retrieve all the interest categories for our list and store them
  // in a global
  .then( function (data) {

    //console.log( 'found ' + data.categories.length + ' categories' );

    g.categories = _.map( data.categories, function( obj ) {
      return { id: obj.id, title: obj.title };
    } );

    //console.log( 'Categories: ', g.categories );

    return g.categories;
  })



  // Read all the categories from MailChimp and store them in g
  .then( function() { return getInterests('Interests'); } )
  .then( function() { return getInterests('Positions'); } )

  // Run thru the list and see what needs updating
  .then( function() { 

    return Promise.map( excelList, function( row ) {
      return updateSubscriber( row );
    }, {concurrency: 1});
 
  })
  .then(function() { console.log("done"); })

  .catch( function( err ) {
    throw new Error( err );
  });

