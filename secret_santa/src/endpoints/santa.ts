import * as ccfapp from "@microsoft/ccf-app";
import { ccf } from "@microsoft/ccf-app/global";

const clientId = "2afc953f-cd9a-42dd-bb68-1f5daef821ac";

export function dump(object) {
  for (var h in object) {
    console.log(h + ": " + object[h]);
  }
}

// Check that the JWT is valid, and call continuation if it is.
export function check_jwt(
  request: ccfapp.Request,
  continuation: (a: ccfapp.JwtAuthnIdentity, params) => ccfapp.Response
): ccfapp.Response {
  console.log("Checking JWT");
  var jwt = request.caller as ccfapp.JwtAuthnIdentity;

  // Check correct issuer
  if (jwt.jwt.keyIssuer != "https://login.microsoftonline.com/common/v2.0/") {
    return { statusCode: 401 };
  }

  var payload = jwt.jwt.payload;

  // Check correct application token.
  if (payload.aud != clientId) {
    return { statusCode: 401 };
  }

  console.log("Checking JWT - Success");
  return continuation(jwt, request.params);
}

// Generate a random id for a secret santa group.
export function random_id() {
  var result = "";
  for (var n = 0; n < 4; n++)
    result += Math.floor(Math.random() * Math.pow(2, 32)).toString(36);
  return result;
}

// Default page for talking to CCF using MSAL to get a JWT token.
// Javascript should contain the custom UI features required for this page
// HTML should contain the corresponding HTML UI elements.
// The Java script supplied should provide an `onLogin` function for
// when the page gets the MSAL response that it has logged in.
function page_template(javascript, html) {
  return `
  <!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, shrink-to-fit=no">
<title>Confidential Secret Santa</title>

<script
  type="text/javascript"
  src="https://alcdn.msauth.net/browser/2.37.0/js/msal-browser.min.js"
  integrity="sha384-DUSOaqAzlZRiZxkDi8hL7hXJDZ+X39ZOAYV9ZDx44gUv9pozmcunJH02tjSFLPnW"
  crossorigin="anonymous"></script>

<!-- adding Bootstrap 4 for UI components  -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css"
  integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">
<link rel="SHORTCUT ICON" href="https://c.s-microsoft.com/favicon.ico?v2" type="image/x-icon">
</head>

<body>
<nav class="navbar navbar-expand-lg navbar-dark bg-primary">
  <a class="navbar-brand" href="/">Confidential Secret Santa</a>
  <div class="btn-group ml-auto dropleft">
    <button id="CreateGroup" class="btn btn-primary" id="createGroup" onclick="createGroup()" hidden="true">
      Create Group
    </button>
    <button type="button" id="SignIn" class="btn btn-secondary" onclick="signIn()">
      Sign In
    </button>
  </div>
</nav>
<br>
${html}

<!-- importing bootstrap.js and supporting js libraries -->
<script src="https://code.jquery.com/jquery-3.4.1.slim.min.js"
  integrity="sha384-J6qa4849blE2+poT4WnyKhv5vZF5SrPo0iEjwBvKU7imGFAV0wwj1yYfoRSJoZ+n"
  crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js"
  integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo"
  crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js"
  integrity="sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6"
  crossorigin="anonymous"></script>

<!-- importing app scripts | load order is important -->
<script type="text/javascript">
  // Config object to be passed to Msal on creation
  const msalConfig = {
      auth: {
          clientId: "${clientId}",
          authority: "https://login.microsoftonline.com/consumers",
          redirectUri: "https://localhost:8000/app/",
      },
      cache: {
          cacheLocation: "sessionStorage", // This configures where your cache will be stored
          storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
      },
      system: {
          allowNativeBroker: false, // Disables WAM Broker
          loggerOptions: {
              loggerCallback: (level, message, containsPii) => {
                  if (containsPii) {
                      return;
                  }
                  switch (level) {
                      case msal.LogLevel.Error:
                          console.error(message);
                          return;
                      case msal.LogLevel.Info:
                          console.info(message);
                          return;
                      case msal.LogLevel.Verbose:
                          console.debug(message);
                          return;
                      case msal.LogLevel.Warning:
                          console.warn(message);
                          return;
                  }
              }
          }
      }
  };

  // Add here scopes for id token to be used at MS Identity Platform endpoints.
  // openid required to get JWT token for talking to CCF.
  const loginRequest = {
      scopes: ["User.Read", "openid"]
  };

  function dump(object) {
    for (var field in object)
      console.log(field + ": " + object[field]);
  }

  let accountId = "";

  /*
  * Create the main myMSALObj instance
  * configuration parameters are located at authConfig.js
  */
  const myMSALObj = new msal.PublicClientApplication(msalConfig); 

  myMSALObj.initialize().then(() => {
      // Register Callbacks for Redirect flow
      myMSALObj.handleRedirectPromise().then(handleResponse).catch((error) => {
          console.log(error);
      });
  });

  const signInButton = document.getElementById("SignIn");

  function handleResponse(resp) {
      if (resp !== null) {
          accountId = resp.account.homeAccountId;
          onLogin();
        } else {
          // need to call getAccount here?
          const currentAccounts = myMSALObj.getAllAccounts();
          if (!currentAccounts || currentAccounts.length < 1) {
              signIn();
          } else if (currentAccounts.length > 1) {
              // Add choose account code here
          } else if (currentAccounts.length === 1) {
              accountId = currentAccounts[0].homeAccountId;
              onLogin();
            }
      }
      signInButton.setAttribute("onclick", "signOut();");
      signInButton.setAttribute("id", "signOutButton");
      signInButton.setAttribute("class", "btn btn-success");
      signInButton.innerHTML = "Sign Out";
  }

  async function signIn() {
      return myMSALObj.loginRedirect(loginRequest);
  }

  function signOut() {
      const logoutRequest = {
          account: myMSALObj.getAccountByHomeId(accountId)
      };
      myMSALObj.logoutRedirect(logoutRequest);
  }

  async function getTokenRedirect(request, account) {
      request.account = account;
      return await myMSALObj.acquireTokenSilent(request).catch(async (error) => {
          console.log("silent token acquisition fails.");
          if (error instanceof msal.InteractionRequiredAuthError) {
              // fallback to interaction when silent call fails
              console.log("acquiring token using redirect");
              myMSALObj.acquireTokenRedirect(request);
          } else {
              console.error(error);
          }
      });
  }

  // Helper function to call CCF API endpoint 
  // using authorization bearer token scheme
  function callCCF(endpoint, method, accessToken, callback) {
      const headers = new Headers();
      const bearer = "Bearer " + accessToken;

      headers.append("Authorization", bearer);

      const options = {
          method: method,
          headers: headers
      };

      console.log("request made to CCF at: " + new Date().toString());

      fetch(endpoint, options)
          .then(response => 
              {
                  console.log("RESPONSE:" + response.status + " on " + endpoint); 
                  return response.json();
              })
          .then(response => callback(response, endpoint))
          .catch(error => console.log(error));
  }

  ${javascript}
  </script>
</body>

</html>`;
}

export function homepage(request: ccfapp.Request): ccfapp.Response {
  console.log("Called: homepage");
  const html = `
<div class="row" style="margin:auto">
  <div id="card-div" class="col-md-3" style="display:none">
    <div class="card text-left">
      <div class="card-body">
        <h5 class="card-title" id="WelcomeMessage">Please sign-in to register and see who you should buy a gift for.</h5>
        <div id="profile-div"><p/></div>
      </div>
    </div>
  </div>
</div>`;

  const javascript = `
  // Select DOM elements to work with
  const welcomeDiv = document.getElementById("WelcomeMessage");
  const cardDiv = document.getElementById("card-div");
  const profileButton = document.getElementById("seeProfile");
  const createGroupButton = document.getElementById("CreateGroup");
  const profileDiv = document.getElementById("profile-div");

  async function updateUI(data) {
    const new_body = document.createElement("p");
    for (var i in data.groups) {
      var group = data.groups[i];
      var groupCard = document.createElement("div");
      groupCard.className = "card text-left";
      var groupP = document.createElement("div");
      groupP.className = "card-body";
      groupCard.appendChild(groupP);

      groupP.innerHTML = "<h5 class='card-title'>" + group.name + "</h5>";
      for (var j in group.members) {
        var memberLI = document.createElement("li");
        var member = group.members[j];
        var memberEntry = memberLI;
        if (member.buying)
        {
          memberEntry = document.createElement("strong");
          memberLI.appendChild(memberEntry);
        }
        memberEntry.innerHTML = member.name + " (" + member.email + ")" + (member.buying ? " <-- buy for" : "");
        groupP.appendChild(memberLI);
      }
      groupP.innerHTML += "<br/><button id='ffee33' class='btn btn-primary' onclick='createLink(this.id)'>Copy Link</button>";
      new_body.appendChild(groupCard);
    }
    cardDiv.appendChild(new_body);
  }

  function createLink(id) {
    navigator.clipboard.writeText("Link:" + id);
  }

  async function seeProfile() {
    let ccfEndpoint = "https://localhost:8000/app/jwt";

    const currentAcc = myMSALObj.getAccountByHomeId(accountId);
    if (currentAcc) {
        const response = await getTokenRedirect(loginRequest, currentAcc).catch(error => {
            console.log(error);
        });
 
        function updateUI(data, endpoint) {
          const name = document.createElement("p");
          name.innerHTML = "<strong>Name: </strong>" + data.displayName + " (" + data.email + ") - " + data.count;
          profileDiv.appendChild(name);
        }
  
        // Using idToken as this is the JWT, don't use AccessToken as this is
        // not acceptable to CCF.
        callCCF(ccfEndpoint, "GET", response.idToken, updateUI);
    }
  }

  async function createGroup() {
    let ccfEndpoint = "https://localhost:8000/app/jwt/create";

    const currentAcc = myMSALObj.getAccountByHomeId(accountId);
    if (currentAcc) {
        const response = await getTokenRedirect(loginRequest, currentAcc).catch(error => {
            console.log(error);
        });
 
        function updateUI(data, endpoint) {
          const name = document.createElement("p");
          name.id = "p-" + data.group;
          name.innerHTML = "<strong>Group: </strong>" + data.group + "<button class='btn btn-primary' id='" + data.group + "' onclick='joinGroup(this.id)'>Join Group</button>";
          profileDiv.appendChild(name);
        }
  
        // Using idToken as this is the JWT, don't use AccessToken as this is
        // not acceptable to CCF.
        callCCF(ccfEndpoint, "POST", response.idToken, updateUI);
    }
  }

  async function joinGroup(group_id)
  {
    let ccfEndpoint = "https://localhost:8000/app/jwt/join/" + group_id;
    const currentAcc = myMSALObj.getAccountByHomeId(accountId);
    if (currentAcc) {
      const response = await getTokenRedirect(loginRequest, currentAcc).catch(error => {
          console.log(error);
      });

      function updateUI(data, endpoint) {
        var button = document.getElementById(group_id);
        button.style.display = "none";

        var paragraph = document.getElementById("p-" + group_id);
        for (var i in data.members)
        {
          const members = document.createElement("p");
          let user = data.members[i];
          members.innerHTML = user.name + " (" + user.email + ")";
          paragraph.appendChild(members);
        }
      }

      // Using idToken as this is the JWT, don't use AccessToken as this is
      // not acceptable to CCF.
      callCCF(ccfEndpoint, "POST", response.idToken, updateUI);
    }
  }

  function onLogin() {
      // Reconfiguring DOM elements
      cardDiv.style.display = "initial";
      welcomeDiv.innerHTML = "Welcome to Secret Santa";

      createGroupButton.hidden = false;
      //seeProfile();
      updateUI(
        {groups: [
          {
            name: "Foo",
            members: [
              {name: "Matt", email: "foo@bar", buying: false},
              {name: "Adam", email: "bar@foo", buying: true},
              {name: "Lisa", email: "cllr@foo", buying: false},
            ] 
          },
          {
            name: "Bar",
            members: []
          }
        ]});
  }`;

  return {
    body: page_template(javascript, html),
    statusCode: 200,
    headers: {
      "content-type": "text/html",
    },
  };
}

let access_count = ccfapp.typedKv("access_count", ccfapp.string, ccfapp.uint32);

export function homepage_jwt(request: ccfapp.Request): ccfapp.Response {
  return check_jwt(request, function (jwt: ccfapp.JwtAuthnIdentity, params) {
    var payload = jwt.jwt.payload;
    var user_id = payload.sub;
    var email = payload.email;
    var displayName = payload.name;

    var count = access_count.has(user_id) ? access_count.get(user_id) : 0;
    access_count.set(user_id, count + 1);

    return {
      body: {
        email: email,
        displayName: displayName,
        count: count,
      },
      statusCode: 200,
      headers: {
        "content-type": "application/json",
      },
    };
  });
}

class Open {}
class Closed {}
type GroupStatus = {
  status: Open | Closed;
  owner: string;
  members: Array<string>;
};

type User = {
  name: string;
  email: string;
  groups: Array<string>;
};

let group_status = ccfapp.typedKv(
  "group_status",
  ccfapp.string,
  ccfapp.json<GroupStatus>()
);

let user_groups = ccfapp.typedKv(
  "user_groups",
  ccfapp.string,
  ccfapp.json<User>()
);

function fresh_group(owner: string): string {
  var id = random_id();
  while (group_status.has(id)) {
    id = random_id();
  }
  group_status.set(id, { status: new Open(), owner: owner, members: [] });
  return id;
}

export function create_jwt(request: ccfapp.Request): ccfapp.Response {
  return check_jwt(request, function (jwt: ccfapp.JwtAuthnIdentity, params) {
    var id = fresh_group(jwt.jwt.payload.sub);
    return {
      body: {
        group: id,
      },
      statusCode: 200,
      headers: {
        "content-type": "application/json",
      },
    };
  });
}

export function join(request: ccfapp.Request): ccfapp.Response {
  return {};
}

function get_or_default_user(
  user_id: string,
  email: string,
  displayName: string
) {
  if (!user_groups.has(user_id)) {
    return { name: displayName, email: email, groups: [] };
  }
  return user_groups.get(user_id);
}

type UserDisplay = { name: string; email: string };

export function join_jwt(request: ccfapp.Request): ccfapp.Response {
  return check_jwt(request, function (jwt: ccfapp.JwtAuthnIdentity, params) {
    var payload = jwt.jwt.payload;
    var user_id = payload.sub;
    var email = payload.email;
    var displayName = payload.name;
    let user = get_or_default_user(user_id, email, displayName);

    let group_id = params.group_id;
    let group = group_status.get(group_id);

    group.members.push(user_id);
    user.groups.push(group_id);

    group_status.set(group_id, group);
    user_groups.set(user_id, user);

    var result = { group: group_id, members: new Array<UserDisplay>() };
    for (var i in group.members) {
      let uid = group.members[i];
      let show_user = uid == user_id ? user : user_groups.get(uid);
      let user_display = { name: show_user.name, email: show_user.email };
      result.members.push(user_display);
    }

    return {
      body: result,
      statusCode: 200,
      headers: {
        "content-type": "application/json",
      },
    };
  });
}
