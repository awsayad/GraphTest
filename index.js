const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

async function main() {
  console.log('JavaScript Graph App-Only Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  const choices = [
    'Display access token',
    'List users',
    'Make a Graph call'
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listUsersAsync();
        break;
      case 2:
        // Run any Graph code
        await makeGraphCallAsync();
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

main();

function initializeGraph(settings) {
    graphHelper.initializeGraphForAppOnlyAuth(settings);
  }
  
  async function displayAccessTokenAsync() {
    try {
      const appOnlyToken = await graphHelper.getAppOnlyTokenAsync();
      console.log(`App-only token: ${appOnlyToken}`);
    } catch (err) {
      console.log(`Error getting app-only access token: ${err}`);
    }
  }
  
  async function listUsersAsync() {
    try {
      const userPage = await graphHelper.getUsersAsync();
      const users = userPage.value;
  
      // Output each user's details
      for (const user of users) {
        console.log(`User: ${user.displayName ?? 'NO NAME'}`);
        console.log(`  ID: ${user.id}`);
        console.log(`  Email: ${user.mail ?? 'NO EMAIL'}`);
        console.log(`  Job title: ${user.jobTitle ?? 'NO Title'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more users
      // available on the server
      const moreAvailable = userPage['@odata.nextLink'] != undefined;
      console.log(`\nMore users available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting users: ${err}`);
    }
  }
  
  async function makeGraphCallAsync() {
    try {
      const groupPage = await graphHelper.makeGraphCallAsync ();
      const groups = groupPage.value;
  
      // Output each group's details
      for (const group of groups) {
        console.log(`Group: ${group.displayName ?? 'NO NAME'}`);
        console.log(`  ID: ${group.id}`);
        console.log(`  Email: ${group.mail ?? 'NO Mail'}`);
        console.log(`  Description: ${group.description ?? 'NO Description'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more users
      // available on the server
      const moreAvailable = groupPage['@odata.nextLink'] != undefined;
      console.log(`\nMore groups available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting groups: ${err}`);
    }
  }

  // To Test the application, open CMD, change directory to app file path (C:\Scripts\JavaScript App-Only Graph Tutorial) and run "node index.js"
