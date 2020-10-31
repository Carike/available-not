// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Select DOM elements to work with
const authenticatedNav = document.getElementById("authenticated-nav");
const accountNav = document.getElementById("account-nav");
const mainContainer = document.getElementById("main-container");

const Views = { error: 1, home: 2, calendar: 3, presence: 4 };

function createElement(type, className, text) {
  var element = document.createElement(type);
  element.className = className;

  if (text) {
    var textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

function showAuthenticatedNav(account, view) {
  authenticatedNav.innerHTML = "";

  if (account) {
    // Add Calendar link
    var calendarNav = createElement("li", "nav-item");

    var calendarLink = createElement(
      "button",
      `btn btn-link nav-link${view === Views.calendar ? " active" : ""}`,
      "Calendar"
    );
    calendarLink.setAttribute("onclick", "getEvents();");
    calendarNav.appendChild(calendarLink);

    authenticatedNav.appendChild(calendarNav);

    // Add Presence link
    var presenceNav = createElement("li", "nav-item");

    var presenceLink = createElement(
      "button",
      `btn btn-link nav-link${view === Views.presence ? " active" : ""}`,
      "Presence"
    );
    presenceLink.setAttribute("onclick", "getPresence();");
    presenceNav.appendChild(presenceLink);

    authenticatedNav.appendChild(presenceNav);
  }
}

function showAccountNav(account, view) {
  accountNav.innerHTML = "";

  if (account) {
    // Show the "signed-in" nav
    accountNav.className = "nav-item dropdown";

    var dropdown = createElement("a", "nav-link dropdown-toggle");
    dropdown.setAttribute("data-toggle", "dropdown");
    dropdown.setAttribute("role", "button");
    accountNav.appendChild(dropdown);

    var userIcon = createElement(
      "i",
      "far fa-user-circle fa-lg rounded-circle align-self-center"
    );
    userIcon.style.width = "32px";
    dropdown.appendChild(userIcon);

    var menu = createElement("div", "dropdown-menu dropdown-menu-right");
    dropdown.appendChild(menu);

    var userName = createElement("h5", "dropdown-item-text mb-0", account.name);
    menu.appendChild(userName);

    var userEmail = createElement(
      "p",
      "dropdown-item-text text-muted mb-0",
      account.userName
    );
    menu.appendChild(userEmail);

    var divider = createElement("div", "dropdown-divider");
    menu.appendChild(divider);

    var signOutButton = createElement("button", "dropdown-item", "Sign out");
    signOutButton.setAttribute("onclick", "signOut();");
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = "nav-item";

    var signInButton = createElement(
      "button",
      "btn btn-link nav-link",
      "Sign in"
    );
    signInButton.setAttribute("onclick", "signIn();");
    accountNav.appendChild(signInButton);
  }
}

function showWelcomeMessage(account) {
  // Create jumbotron
  var jumbotron = createElement("div", "jumbotron");

  var heading = createElement("h1", null, "Available/Not");
  jumbotron.appendChild(heading);

  var lead = createElement(
    "p",
    "lead",
    "Simple app using Microsoft technologies and IoT to show " +
      " others whether you're available or not"
  );
  jumbotron.appendChild(lead);

  if (account) {
    // Welcome the user by name
    var welcomeMessage = createElement("h4", null, `Welcome ${account.name}!`);
    jumbotron.appendChild(welcomeMessage);

    var callToAction = createElement(
      "p",
      null,
      "Use the navigation bar at the top of the page to get started."
    );
    jumbotron.appendChild(callToAction);
  } else {
    // Show a sign in button in the jumbotron
    var signInButton = createElement(
      "button",
      "btn btn-primary btn-large",
      "Click here to sign in"
    );
    signInButton.setAttribute("onclick", "signIn();");
    jumbotron.appendChild(signInButton);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(jumbotron);
}

function showError(error) {
  var alert = createElement("div", "alert alert-danger");

  var message = createElement("p", "mb-3", error.message);
  alert.appendChild(message);

  if (error.debug) {
    var pre = createElement("pre", "alert-pre border bg-light p-2");
    alert.appendChild(pre);

    var code = createElement(
      "code",
      "text-break text-wrap",
      JSON.stringify(error.debug, null, 2)
    );
    pre.appendChild(code);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(alert);
}

// <showCalendar>
function showCalendar(events) {
  var div = document.createElement("div");

  div.appendChild(createElement("h1", null, "Calendar"));

  var table = createElement("table", "table");
  div.appendChild(table);

  var thead = document.createElement("thead");
  table.appendChild(thead);

  var headerrow = document.createElement("tr");
  thead.appendChild(headerrow);

  var organizer = createElement("th", null, "Organizer");
  organizer.setAttribute("scope", "col");
  headerrow.appendChild(organizer);

  var subject = createElement("th", null, "Subject");
  subject.setAttribute("scope", "col");
  headerrow.appendChild(subject);

  var start = createElement("th", null, "Start");
  start.setAttribute("scope", "col");
  headerrow.appendChild(start);

  var end = createElement("th", null, "End");
  end.setAttribute("scope", "col");
  headerrow.appendChild(end);

  var tbody = document.createElement("tbody");
  table.appendChild(tbody);

  for (const event of events.value) {
    var eventrow = document.createElement("tr");
    eventrow.setAttribute("key", event.id);
    tbody.appendChild(eventrow);

    var organizercell = createElement(
      "td",
      null,
      event.organizer.emailAddress.name
    );
    eventrow.appendChild(organizercell);

    var subjectcell = createElement("td", null, event.subject);
    eventrow.appendChild(subjectcell);

    var startcell = createElement(
      "td",
      null,
      moment.utc(event.start.dateTime).local().format("M/D/YY h:mm A")
    );
    eventrow.appendChild(startcell);

    var endcell = createElement(
      "td",
      null,
      moment.utc(event.end.dateTime).local().format("M/D/YY h:mm A")
    );
    eventrow.appendChild(endcell);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(div);
}
// </showCalendar>

// <showPresence>
function showPresence(teamsStatus) {
  // Heading
  var div = document.createElement("div");
  div.appendChild(createElement("h1", null, "Teams presence"));

  // Availability Section
  var availabilityValue = teamsStatus.availability;
  var availabilitySection = document.createElement("p");
  availabilitySection.innerText = "Availability : " + availabilityValue;
  div.appendChild(availabilitySection);

  // Activity Section
  var activityValue = teamsStatus.activity;
  var activitySection = document.createElement("p");
  activitySection.innerText = "Activity : " + activityValue;
  div.appendChild(activitySection);

  var colourLight;
  switch (availabilityValue) {
    case "Available":
    case "AvailableIdle":
      colourLight = "#008000"; // green
      break;

    case "Busy":
    case "BusyIdle":
    case "DoNotDisturb":
      colourLight = "#ff0000"; // red
      break;

    case "Away":
    case "BeRightBack":
      colourLight = "#ffff00"; // yellow
      break;

    case "PresenceUnknown":
      colourLight = "#808080"; // grey
      break;

    case "Offline":
      colourLight = "#800080"; // purple
      break;

    default:
      colourLight = "#ffc0cb"; // pink
  }
  console.log(colourLight);

  var lightbulb = document.createElement("div");
  lightbulb.innerHTML =
    `<div style="width: 56px; height: 56px;">
<svg version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px"
	 viewBox="0 0 512 512" style="enable-background:new 0 0 512 512;" xml:space="preserve">
<g>
	<g>
		<path style="fill:` +
    colourLight +
    `;" d="M256,0C157.885,0,78.063,79.822,78.063,177.936c0,46.98,18.144,91.296,51.088,124.784l2.627,2.664
			c12.84,13.511,22.519,29.458,28.548,46.871l-0.592,0.372c-11.394,7.163-18.198,19.455-18.198,32.875
			c0,13.279,6.624,25.167,17.105,32.202c-0.923,3.334-1.422,6.809-1.422,10.357c0,21.417,17.424,38.841,38.841,38.841h3.793
			C205.545,492.665,228.551,512,256,512s50.455-19.336,56.147-45.099h3.793c21.417,0,38.841-17.424,38.841-38.841
			c0-3.548-0.499-7.023-1.422-10.357c10.482-7.034,17.105-18.923,17.105-32.202c0-13.42-6.802-25.711-18.198-32.875l-0.592-0.372
			c6.005-17.348,15.634-33.24,28.405-46.719l2.77-2.816c32.945-33.487,51.088-77.803,51.088-124.784C433.937,79.822,354.115,0,256,0
			z M256,481.583c-10.465,0-19.562-5.964-24.074-14.671h48.147C275.563,475.619,266.466,481.583,256,481.583z M315.94,436.483
			h-2.429v-0.105l-15.313,0.105h-84.371l-15.337-0.205v0.205h-2.429c-4.645,0-8.424-3.779-8.424-8.424
			c0-1.561,0.433-3.064,1.22-4.363h134.287c0.787,1.298,1.22,2.801,1.22,4.363C324.364,432.705,320.585,436.483,315.94,436.483z
			 M334.822,393.281h-0.966H178.145h-0.966c-3.133-1.27-5.226-4.306-5.226-7.779c0-1.745,0.538-3.405,1.51-4.792h165.075
			c0.971,1.386,1.51,3.047,1.51,4.792C340.048,388.974,337.955,392.01,334.822,393.281z M239.773,350.293v-90.148h32.453v90.148
			H239.773z M362.011,280.522l-0.749,0.747c-19.394,19.326-33.376,42.992-40.879,69.024h-17.738v-90.148h9.559
			c22.043,0,39.976-17.934,39.976-39.977c0-22.042-17.934-39.976-39.976-39.976c-22.042,0-39.976,17.934-39.976,39.976v9.56h-32.453
			v-9.56c0-22.042-17.934-39.976-39.976-39.976s-39.976,17.934-39.976,39.976c0,22.043,17.934,39.977,39.976,39.977h9.559v90.148
			h-17.738c-7.503-26.033-21.486-49.699-40.88-69.024l-0.71-0.706c-26.802-27.663-41.549-64.064-41.549-102.628
			c0-81.342,66.177-147.518,147.519-147.518S403.52,96.594,403.52,177.936C403.52,216.481,388.787,252.864,362.011,280.522z
			 M302.644,229.728v-9.56c0-5.27,4.288-9.559,9.559-9.559s9.559,4.288,9.559,9.559s-4.288,9.56-9.559,9.56H302.644z
			 M209.356,220.168v9.56h-9.559c-5.271,0-9.559-4.288-9.559-9.56c0-5.27,4.288-9.559,9.559-9.559S209.356,214.897,209.356,220.168z
			"/>
	</g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
<g>
</g>
</svg>
</div>
`;

  div.appendChild(lightbulb);

  mainContainer.innerHTML = "";
  mainContainer.appendChild(div);
}
// </showPresence>

// <updatePage>
function updatePage(account, view, data) {
  if (!view || !account) {
    view = Views.home;
  }

  showAccountNav(account);
  showAuthenticatedNav(account, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(account);
      break;
    case Views.calendar:
      showCalendar(data);
      break;
    case Views.team:
      showCalendar(data);
      break;
    case Views.presence:
      showPresence(data);
      break;
  }
}
// </updatePage>

updatePage(null, Views.home);
