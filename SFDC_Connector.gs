/**
* @author       Nicolas Robert
* @date         20/07/2022
* @description  Google Spreadsheets Script to query Salesforce API data based on https://github.com/tinybigideas/Google-Cloud-Connecter
* @URL          https://github.com/DarkPampers/SFDC-Connector/edit/master/CloudConnector.gs
*/

var userProperties = PropertiesService.getUserProperties();
var USERNAME_PROPERTY_NAME = "username";
var PASSWORD_PROPERTY_NAME = "password";
var SECURITY_TOKEN_PROPERTY_NAME = "securityToken";
var SESSION_ID_PROPERTY_NAME = "sessionId";
var SERVICE_URL_PROPERTY_NAME = "serviceUrl";
var INSTANCE_URL_PROPERTY_NAME = "instanceUrl";
var IS_SANDBOX_PROPERTY_NAME = "isSandbox";
var NEXT_RECORDS_URL_PROPERTY_NAME = "nextRecordsUrl";
var SOBJECT_ATTRIBUTES_PROPERTY_NAME = "sObjectAttributes";
var SANDBOX_SOAP_URL = "https://test.salesforce.com/services/Soap/u/30.0";
var PRODUCTION_SOAP_URL = "https://login.salesforce.com/services/Soap/u/30.0";

/**
 * @return String Username.
 */
function getUsername() {
  var key = userProperties.getProperty(USERNAME_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
};

/**
 * @param String Username.
 */
function setUsername(key) {
  userProperties.setProperty(USERNAME_PROPERTY_NAME, key);
};

/**
 * @return String Password.
 */
function getPassword() {
  var key = userProperties.getProperty(PASSWORD_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
};

/**
 * @param String Password.
 */
function setPassword(key) {
  userProperties.setProperty(PASSWORD_PROPERTY_NAME, key);
};

/**
 * @return String Security Token.
 */
function getSecurityToken() {
  var key = userProperties.getProperty(SECURITY_TOKEN_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
};

/**
 * @param String Security Token.
 */
function setSecurityToken(key) {
  userProperties.setProperty(SECURITY_TOKEN_PROPERTY_NAME, key);
}

/**
 * @return String Session Id.
 */
function getSessionId() {
  var key = userProperties.getProperty(SESSION_ID_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
};

/**
 * @param String Session Id.
 */
function setSessionId(key) {
  userProperties.setProperty(SESSION_ID_PROPERTY_NAME, key);
};

/**
 * @return String Instance URL.
 */
function getInstanceUrl() {
  var key = userProperties.getProperty(INSTANCE_URL_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
};

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
  userProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
};

/**
 * @param String use sandbox url.
 */
function setUseSandbox(key) {
  userProperties.setProperty(IS_SANDBOX_PROPERTY_NAME, key);
};

/**
 * @return bool if using sandbox.
 */
function getUseSandbox() {
  var key = userProperties.getProperty(IS_SANDBOX_PROPERTY_NAME);
  if (key == null) {
    key = false;
  }
  return key;
};

/**
 * @param String url for next records url.
 */
function setNextRecordsUrl(key) {
  if (key == undefined) {
    key = "";
  }
  userProperties.setProperty(NEXT_RECORDS_URL_PROPERTY_NAME, key);
};

/**
 * @return String url for next records url (querymore).
 */
function getNextRecordsUrl() {
  var key = userProperties.getProperty(NEXT_RECORDS_URL_PROPERTY_NAME);
  if (key == null || key == undefined) {
    key = "";
  }
  return key;
};

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
  userProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
};

/**
 * @return bool if using sandbox.
 */
function getSfdcSoapEndpoint() {
  var isSandbox = getUseSandbox() == "true" ? true : false;
  if (isSandbox) {
    return SANDBOX_SOAP_URL;
  }
  else {
    return PRODUCTION_SOAP_URL;
  }
};

function getRestEndpoint() {
  var queryEndpoint = ".my.salesforce.com";
  var endpoint = getInstanceUrl().replace("api-", "").match("https://[a-z0-9]*");
  return endpoint + queryEndpoint;
};

/**
 * Login script
 */
function login() {
  var utilisateur = getUsername();
  var pass = getPassword();
  var sec = getSecurityToken();
  //Logger.log (utilisateur + " / " +  pass + " / " + sec)
  var message = "<?xml version='1.0' encoding='utf-8'?>"
    + "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' "
    + "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://"
    + "www.w3.org/2001/XMLSchema'>"
    + "<soap:Body>"
    + "<login xmlns='urn:partner.soap.sforce.com'>"
    + "<username>" + utilisateur + "</username>"
    + "<password>" + pass + sec + "</password>"
    + "</login>"
    + "</soap:Body>"
    + "</soap:Envelope>";
  var httpheaders = { SOAPAction: "login" };
  var parameters = {
    muteHttpExceptions: true,
    method: "POST",
    contentType: "text/xml",
    headers: httpheaders,
    payload: message
  };
  try {
    Logger.log("[SFDC] to ["+utilisateur+"] : What are you doing step user ?!");
    var result = UrlFetchApp.fetch(getSfdcSoapEndpoint(), parameters).getContentText();
    var soapResult = XmlService.parse(result);
    var ns1 = XmlService.getNamespace('soapenv', 'http://schemas.xmlsoap.org/soap/envelope/');
    var ns2 = XmlService.getNamespace('loginresponse', 'urn:partner.soap.sforce.com');
    var nodeIds = soapResult.getRootElement().getChild('Body', ns1).getChild('loginResponse', ns2).getChild('result', ns2);
    var sessionId = nodeIds.getChild('sessionId', ns2).getValue();
    var serverUrl = nodeIds.getChild('serverUrl', ns2).getValue();
    setSessionId(sessionId);
    setInstanceUrl(serverUrl);
    Logger.log("[SFDC] to ["+utilisateur+"] : You Did It. The Crazy Son of a Bitch, You Did It !");
    Logger.log("["+utilisateur+"] to [SFDC] : Jokes on you ! I'm into that shit.");
  }
  catch (e) {
    Logger.log("[SFDC] to ["+utilisateur+"] : Ah Ah Ah ! " + e);
    mailRecapEchec();
  }
};

function ReportQueryReplace(reportId) {
  try {
    var results = fetch(getRestEndpoint() + "/services/data/v36.0/analytics/reports/" + reportId + "?includeDetails=true");
    let resultsFormated = RenderResult(results);
    return resultsFormated;
  }
  catch (e) {
    Logger.log(e);
  }
};

function ReportQueryAdd(reportId) {
  try {
    var results = fetch(getRestEndpoint() + "/services/data/v36.0/analytics/reports/" + reportId + "?includeDetails=true");
    Logger.log("Report ID : " + reportId + " Fetched !");
    let resultsFormated = RenderResult(results);
    resultsFormated.shift();
    return resultsFormated;
  }
  catch (e) {
    Logger.log(e);
  }
};

function RenderResult(result) {
  var queryResult = JSON.parse(result);
  var answer = queryResult.factMap["T!T"].rows;
  var headers = queryResult.reportExtendedMetadata.detailColumnInfo;
  var headname = queryResult.reportMetadata.detailColumns;
  var myArray = [];
  var tempArray = [];
  for (i = 0; i < headname.length; i++) {
    tempArray.push(headers[headname[i]].label);
  }
  myArray.push(tempArray);
  for (i = 0; i < answer.length; i++) {
    var tempArray = [];
    function getData(element, index, array) {
      tempArray.push(array[index].label)
    }
    answer[i].dataCells.forEach(getData);
    myArray.push(tempArray);
  }
  for (var i in myArray) {
    for (var j in myArray[i]) {
      if (myArray[i][j] === "-") {
        myArray[i][j] = "";
      }
    }
  }
  Logger.log("SFDC Data Formated ! - length : " + myArray.length)
  return myArray;
}

/**
 * Run SOQL Query in spreadsheet
 */
function SOQLQuery(SOQL) {
  var results = fetch(getRestEndpoint() + "/services/data/v36.0/" + "query?q=" + encodeURIComponent(SOQL));
  //return Utilities.jsonParse(results);
  console.log(getRestEndpoint() + "/services/data/v36.0/" + "query?q=" + encodeURIComponent(SOQL))
  console.log(results)
  //return RenderResult(results);
  return renderGridData(Utilities.jsonParse(results));
};

function ReportToSOQLQuery(reportId) {
  try {
    let results = fetch(getRestEndpoint() + "/services/data/v36.0/analytics/reports/" + reportId + "?includeDetails=true");
    console.log("Report ID : " + reportId + " Fetched !");
    let queryResult = JSON.parse(results);
    let fields = queryResult.reportMetadata.detailColumns.join(",  ");
    let table = queryResult.reportMetadata.reportType.type;
    let dateFilter = queryResult.reportMetadata.standardDateFilter.column + " >= " + queryResult.reportMetadata.standardDateFilter.startDate + " AND " + queryResult.reportMetadata.standardDateFilter.column + " < " + queryResult.reportMetadata.standardDateFilter.endDate;
    let filter = transfoFilter(queryResult.reportMetadata.reportFilters);
    let requete = " SELECT  " + fields + "  FROM  " + table + "  WHERE  "  + filter + " AND " + dateFilter;
    return requete
  }
  catch (e) {
    console.log(e);
  }
};

function transfoFilter(array) {
  let result = "";
  array.forEach(e => {result += e.column + " " + transfoSigne(e.operator) + " " + e.value + " AND "});
  return result.slice(0, result.length-5);
}

function transfoSigne(string) {
  switch (string) {
    case 'greaterThan':
      return ">";
    case 'greaterOrEqual':
      return ">=";
    case 'lessThan':
      return "<";
    case 'lessOrEqual':
      return "<=";
    case 'equals':
      return "=";
    default:
      console.log(`Sorry, we are out of ${string}.`);
  }
}

/**
 * Get data from API
 */
function fetch(url) {
  
  var httpheaders = { Authorization: "OAuth " + getSessionId() };
  var parameters = { headers: httpheaders, muteHttpExceptions: true };
  return UrlFetchApp.fetch(url, parameters).getContentText();
};

