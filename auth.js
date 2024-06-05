async function initMsal(){
  const msalConfig = {
      auth: {
        clientId: "f851df94-4e6f-42f7-8524-dd9af747a3e1",
        authority: "https://login.microsoftonline.com/organizations/",
        redirectUri: "https://start.sampol.com/",
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
      },
    };
 const msalInstance = new msal.PublicClientApplication(msalConfig);
 
 return msalInstance;
 }
 
 async function login(num) {
  hideSampolInfo();
  const msalInstance = await initMsal();

  const loginRequest = { scopes: ["openid", "profile", "User.Read", "Calendars.read"] };

  let accounts = await msalInstance.getAllAccounts();
  let auth = "";
  let tokenExpired = sessionStorage.getItem("tokenExpired");

  if (accounts.length > 0 && !tokenExpired) {
    hideEntra();
    try {
      auth = await silentLogin(msalInstance, loginRequest, accounts[0]);
      sessionStorage.removeItem("tokenExpired");

      showSampolInfo(accounts[0].name);
    } catch (error) {
      sessionStorage.setItem("tokenExpired", true);
      document.getElementById('entra').style = "display: block";
    }

  } else {
    if (num == 1) {
      hideEntra()
      auth = await showPopup(msalInstance, loginRequest);
      sessionStorage.removeItem("tokenExpired");
      accounts = await msalInstance.getAllAccounts();
      showSampolInfo(accounts[0].name);
    } else {
      document.getElementById('entra').style = "display: block";
    }

  }

  displayEvents(auth);
}

async function showPopup(msal, loginRequest) {
  const authResult = await msal.loginPopup(loginRequest);
  return auth = "Bearer " + authResult.accessToken;
}

async function silentLogin(msal, loginRequest, account) {

  const silentResult = await msal.acquireTokenSilent({
    scopes: loginRequest.scopes,
    account: account
  });
  auth = "Bearer " + silentResult.accessToken;
  return auth;
}

async function displayEvents(auth) {
  const events = await getEvents(auth);
  if (events.length > 0) {

    sortByDate(events);
    updateTimeZone(events);
    eventos = events.slice(0, 4);
    eventos.forEach((event, idx, arr) => {
      var url = "";
      if (event.onlineMeeting != null && event.onlineMeeting.joinUrl != null) {
        url = event.onlineMeeting.joinUrl;
      } else {
        url = event.webLink;
      }
      var start = new Date(event.start.dateTime);
      var end = new Date(event.end.dateTime);
      const optionsDay = {
        month: 'numeric',
        day: 'numeric',
      };
      const optionsHour = { hour: '2-digit', minute: '2-digit' };
      if (idx === arr.length - 1) {
        var templateString =
          '<a style="text-decoration: none; color: white !important;" class="card" href=' + url + ' target="_blank">' +
          '<div class="cardInfo">' +
          '<p  class="card-title" style="margin-right: 5px;margin-top: auto;margin-bottom: auto;width: 296px;">' + event.subject + '</p>' +
          '</div>' +
          "<p style='text-align: justify;'>" + event.location.displayName + "</p>" +
          '<p class="date-card" style="margin-bottom: 5px;border-bottom: 1px solid white;text-align: justify;"> Desde ' + start.toLocaleTimeString('es-ES', optionsHour) + ' ' + start.toLocaleDateString('es-ES', optionsDay) + ' hasta ' + end.toLocaleTimeString('es-ES', optionsHour) + ' ' + end.toLocaleDateString('es-ES', optionsDay) + '</p>';
        '</a>';
      } else {
        var templateString =

          '<a style="text-decoration: none; color: white !important;" class="card"href=' + url + ' target="_blank">' +
          '<div class="cardInfo">' +
          '<p  class="card-title" style="margin-right: 5px; margin-top: auto;margin-bottom: auto;">' + event.subject + '</p>' +
          "<p>" + event.location.displayName + "</p>" +
          '</div>' +
          '<p class="date-card" style="margin-bottom: 5px;border-bottom: 1px solid white"> Desde ' + start.toLocaleTimeString('es-ES', optionsHour) + ' ' + start.toLocaleDateString('es-ES', optionsDay) + ' hasta ' + end.toLocaleTimeString('es-ES', optionsHour) + ' ' + end.toLocaleDateString('es-ES', optionsDay) + '</p>';
        '</a>';
      }

      var cards = document.getElementById("cards")
      var cardtmp = document.createElement("li");
      if (idx === arr.length - 1) { cardtmp.className = ' fourth'; }
      var card = cards.appendChild(cardtmp);
      card.innerHTML = templateString;
    });
  }
  hideEntra();
  checkOverflow();
}

async function getEvents(auth) {
  const now = new Date();
  const weekAfter = new Date();
  weekAfter.setDate(weekAfter.getDate() + 7);
  const url = "https://graph.microsoft.com/v1.0/me/calendarView";
  params = {
    startDateTime: now.toISOString(),
    endDateTime: weekAfter.toISOString(),
  };
  response = await axios
    .get(url, { headers: { Authorization: auth }, params: params })
    .then(function (respuesta) {
      return respuesta.data.value;
    });
  return response;

}

function hideEntra() {
  document.getElementById("entra").style = "display: none";
}

function sortByDate(events) {
  events.sort(function (a, b) {

    let dateA = new Date(a.start.dateTime);
    let dateB = new Date(b.start.dateTime);


    return dateA - dateB;
  });
}

function updateTimeZone(events) {
  events.forEach(event => {
    let start = new Date(event.start.dateTime);
    let end = new Date(event.end.dateTime);

    const dt = new Date();
    let diffTZ = ((dt.getTimezoneOffset() / 60) * -1);

    start.setHours(start.getHours() + 1);
    end.setHours(end.getHours() + 1);

    event.start.dateTime = start.toISOString();
    event.end.dateTime = end.toISOString();

  });
}

function checkOverflow() {
  var paragraphs = document.querySelectorAll('p');

  paragraphs.forEach(function (paragraph) {
    if (isTextOverflowed(paragraph)) {
      paragraph.classList.add('overflowed');
    }
  });
}

function isTextOverflowed(element) {
  return element.scrollHeight > element.clientHeight || element.scrollWidth > element.clientWidth;
}

function hideSampolInfo(){
  document.getElementById('leftMenu').classList.add('hide');
  document.getElementById('rightMenu').classList.add('hide');
  document.getElementById('corpmsg').classList.add('hide');
}
function showSampolInfo(userName){
  console.log(userName);
  document.getElementById('leftMenu').classList.remove('hide');
  document.getElementById('rightMenu').classList.remove('hide');
  document.getElementById('corpmsg').classList.remove('hide');
  document.getElementById('userName').innerHTML = userName;
}
