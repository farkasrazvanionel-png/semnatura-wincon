/*
 * Semnatura WINCON - add-in Outlook (event-based activation).
 * La fiecare email nou, raspuns sau redirectionare pune semnatura standard WINCON.
 * Datele colegilor vin din users.json, generat din Microsoft 365 cu scripts/export-users.ps1.
 */

// Adresa de unde se incarca users.json. Se inlocuieste cu domeniul pe care e gazduit add-in-ul.
var DATA_URL = "https://farkasrazvanionel-png.github.io/semnatura-wincon/users.json";

var COMPANY = "WINCON S.R.L.";
var WEBSITE = "www.wincon.ro";
// Logo gazduit (link, nu atasament), ca emailurile sa nu para ca au fisiere atasate.
var LOGO_URL = "https://farkasrazvanionel-png.github.io/semnatura-wincon/wincon-logo.png";
// "middle" = intre functie si contact; "bottom" = sub adresa.
var LOGO_POSITION = "middle";

// Biroul din Microsoft 365 (campul Office) -> adresa din semnatura.
var OFFICES = {
  "bucuresti": "Str. Putul lui Zamfir 32-34, et. 3, Sector 1, Bucuresti",
  "bucharest": "Str. Putul lui Zamfir 32-34, et. 3, Sector 1, Bucuresti",
  "cluj-napoca": "Str. G-ral Ion Dragalina 93, Cluj-Napoca",
  "cluj": "Str. G-ral Ion Dragalina 93, Cluj-Napoca"
};

var STYLE = "font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#000000;";

// Fara diacritice: ghidul WINCON cere text fara diacritice (inclusiv ș/ț cu virgula).
function stripDiacritics(text) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Orice format (0770915249, +4 0733101747, +40 (774) 612 028) -> +40 7XX XXX XXX.
function formatPhone(raw) {
  var digits = String(raw || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.indexOf("0040") === 0) digits = digits.slice(2);
  if (digits.length === 10 && digits[0] === "0") digits = "4" + digits;
  if (digits.length === 11 && digits.indexOf("40") === 0) {
    return "+40 " + digits.slice(2, 5) + " " + digits.slice(5, 8) + " " + digits.slice(8);
  }
  return "+" + digits;
}

function officeAddress(office) {
  var key = stripDiacritics(office).toLowerCase();
  return OFFICES[key] || stripDiacritics(office);
}

function normalizeUser(user, fallbackName, fallbackEmail) {
  user = user || {};
  return {
    name: stripDiacritics(user.name || fallbackName),
    title: stripDiacritics(user.title),
    phone: formatPhone(user.mobile),
    email: String(user.email || fallbackEmail || "").toLowerCase(),
    address: officeAddress(user.office)
  };
}

// compact = true pentru raspunsuri si redirectionari (fara logo, ca raspunsurile sa ramana usoare).
function buildSignature(u, compact, logoUrl, logoPosition) {
  logoUrl = logoUrl || LOGO_URL;
  logoPosition = logoPosition || LOGO_POSITION;
  var name = "<b>" + escapeHtml(u.name) + "</b>";
  var phoneLink = u.phone
    ? '<a href="tel:' + u.phone.replace(/ /g, "") + '" style="color:#000000;text-decoration:none;">' + u.phone + "</a>"
    : "";

  if (compact) {
    var line1 = [name, escapeHtml(u.title), COMPANY].filter(Boolean).join(" | ");
    return '<div style="' + STYLE + '"><p style="margin:0;">' + line1 +
      (phoneLink ? "<br>m: " + phoneLink : "") + "</p></div>";
  }

  var who = [name, escapeHtml(u.title)].filter(Boolean).join("<br>");
  var logo = '<a href="https://' + WEBSITE + '"><img src="' + logoUrl + '" alt="' + COMPANY +
    '" width="180" height="42" style="display:block;border:0;width:180px;height:42px;"></a>';
  var contact = [];
  if (phoneLink) contact.push("m: " + phoneLink);
  if (u.email) contact.push('e: <a href="mailto:' + escapeHtml(u.email) + '">' + escapeHtml(u.email) + "</a>");
  contact.push('w: <a href="https://' + WEBSITE + '">' + WEBSITE + "</a>");
  if (u.address) contact.push(escapeHtml(u.address));

  return '<div style="' + STYLE + '">' +
    '<p style="margin:0;">Cu stima | Kind regards</p><br>' +
    '<p style="margin:0;">' + who + "</p>" +
    (logoPosition === "bottom"
      ? '<br><p style="margin:0;">' + contact.join("<br>") + "</p>" +
        '<p style="margin:10px 0 0 0;">' + logo + "</p>"
      : '<p style="margin:8px 0;">' + logo + "</p>" +
        '<p style="margin:0;">' + contact.join("<br>") + "</p>") +
    "</div>";
}

// Outlook opreste add-in-ul dupa 60 de secunde; nimic nu are voie sa astepte la nesfarsit.
var WAIT_MS = 4000;
function withTimeout(promise, fallback) {
  return Promise.race([promise, new Promise(function (resolve) { setTimeout(function () { resolve(fallback); }, WAIT_MS); })]);
}

function loadUser(email) {
  return fetch(DATA_URL, { cache: "no-cache" })
    .then(function (r) { return r.ok ? r.json() : {}; })
    .then(function (users) { return users[String(email).toLowerCase()]; })
    .catch(function () { return undefined; });
}

function getComposeType(item) {
  return new Promise(function (resolve) {
    if (!item.getComposeTypeAsync) return resolve("newMail");
    item.getComposeTypeAsync(function (result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value.composeType : "newMail");
    });
  });
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;

  Promise.all([withTimeout(getComposeType(item), "newMail"), withTimeout(loadUser(profile.emailAddress), undefined)])
    .then(function (results) {
      var compact = results[0] !== "newMail";
      var html = buildSignature(normalizeUser(results[1], profile.displayName, profile.emailAddress), compact);
      if (!item.body.setSignatureAsync) return event.completed();
      var setSignature = function () {
        item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (result) {
          if (result && result.status !== Office.AsyncResultStatus.Succeeded) console.error("Semnatura WINCON:", result.error);
          event.completed();
        });
      };
      // Opreste semnatura salvata manual in Outlook, ca sa nu apara doua semnaturi.
      // Pe telefon functia poate lipsi; acolo semnatura aplicatiei se opreste din setarile Outlook.
      if (item.disableClientSignatureAsync) item.disableClientSignatureAsync(setSignature);
      else setSignature();
    })
    .catch(function (err) { console.error("Semnatura WINCON:", err); event.completed(); });
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}

if (typeof module !== "undefined") {
  module.exports = { stripDiacritics: stripDiacritics, formatPhone: formatPhone, normalizeUser: normalizeUser, buildSignature: buildSignature };
}
