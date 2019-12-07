function refresher() {  
  var spreadsheetUrls = [
     
]; 
  
  for(var s = 0; s < spreadsheetUrls.length; s++) {
    Logger.log("Обрабатываю: " + spreadsheetUrls[s]);
    processAllScripts(spreadsheetUrls[s]); //Все сотрудники обрабатываются по очереди. Если один накосячил с данными в таблице - обновление не получат он и все кто после него
  }
}

function processAllScripts(spreadsheetUrl){
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Yandex'));
  
  var loginsCol = sheet.getRange('E2:E').getValues().filter(String); // логины, нельзя пропускать строки
  var startDateCol = sheet.getRange('I2:I').getValues().filter(String); //дата начала
  var goalsCol = sheet.getRange('A2:A' + (loginsCol.length + 1)).getValues(); //цели
  var resultCol = sheet.getRange('B2:B').getValues(); //результаты
  var flagsCol = sheet.getRange('C2:C' + (loginsCol.length + 1)).getValues(); //флаги
  var clientNameCol = sheet.getRange('D2:D' + (loginsCol.length + 1)).getValues(); //имена клиентов
  var countersCol = sheet.getRange('H2:H' + (loginsCol.length + 1)).getValues(); //ссылки на счетчики
    
  trimArray(loginsCol); //убираем лишние пробелы, если они есть
  
  if(loginsCol.length > 0) {
    Logger.log('Проверяю имена клиентов');
    checkClientName(clientNameCol, loginsCol, sheet);
    
    Logger.log('Проверяю список целей Метрики');
    goalsCol = checkGoals(goalsCol, loginsCol, sheet);
    
    Logger.log('Проверяю счетчиики Метрики');
    checkCounterLinks(countersCol, loginsCol, sheet);
    
    Logger.log('Приступаю к отчетам');
    var stats = processReports(loginsCol, startDateCol, goalsCol); //Получаем клики, расход и время последнего обновления
    
    Logger.log('Перехожу к общим счетам');
    var budgets = processBudgets(loginsCol, flagsCol); //Получаем остаток общего счета
    
    Logger.log('Форматирую результаты');    
    var result = formResultStrings(stats, budgets, goalsCol, resultCol); //Собираем все данные одну строку с нужным форматом
    
    Logger.log('Обновляю таблицу'); 
    sheet.getRange('B2:B' + (loginsCol.length + 1)).setValues(result); //Обновление данных в таблице
  } else {
    Logger.log("Пустой список логинов клиентов");
  }
}
//Собираем отчет
function processReports(loginsArray, startDateArray, goalsCol){
  var statsRequests = setStatsRequests(loginsArray, startDateArray, goalsCol); //Формируем запросы к API для каждого логина
  var stats = getStats(statsRequests);  //Отправляем запросы и получаем данные
  return stats;
}
//Собираем данные общего счета
function processBudgets(loginsArray, flagsArray){
  var budgetsLogins = sortLoginsForBudgets(loginsArray, flagsArray); //Группируем логины по агентствам
  var budgetsRequests = setBudgetsRequests(budgetsLogins); //Формируем запросы к API для агентства
  var budgets = getBudgets(budgetsRequests);  //Отправляем запросы и получаем данные
  var budgetsIndexed = indexBudgets(loginsArray, budgets); //Распределяем полученные данные по нужным строкам
  
  return budgetsIndexed;
}
//Избавляем массив от лишних пробелов
function trimArray(array){  
  for(var i = 0; i < array.length; i++)
    array[i][0] = array[i][0].trim();
}
//Выбор токена для агентства
function chooseToken(clientLogin){  
  var tokenBY = 'qq';
  var tokenRU = 'qq';
  var tokenKZ = 'qq';
  
  if(clientLogin.match(/ru/) != null) return tokenRU;
    else if(clientLogin.match(/by/) != null) return tokenBY;
    else if(clientLogin.match(/kz/i) != null) return tokenKZ;
}

//Составляем строку формата 'Параметр: Значение'
function formResultStrings(stats, budgets, goalsCol, resCol) {
  var returnArray = [];
  for(var j = 0; j < stats.length; j++) {
    if(stats[j] === 'Skipped') {
      returnArray[j] = resCol[j]; //Если отчет пропущен - оставляем прежнее значение ячейки
    } else {
      var delimiter = '/|/';
      var splitedString = stats[j].split(/\n/); //Когда все хорошо, получаем массив вида [Дата обновления, Заголовки отчета, Значения отчета]
      var headersArray = splitedString[1].split(/\t/); //Разделяем заголовки [Клики, Расход]
      var statsArray = splitedString[2].split(/\t/); //Делим значения [Клики, Расход]
      var lastReportDate = 'Date: ' + Utilities.formatDate(new Date(splitedString[0]), 'GMT+3', 'HH:mm dd.MM.yyyy') + delimiter; //Форматируем дату обновления
      if(!(goalsCol[j][0] === 'Ошибка' || goalsCol[j][0] === '')){
        var splitedGoalsString = goalsCol[j][0].split(delimiter);
        var goalsIDs = splitedGoalsString[0].split(',');
        var goalsNames = splitedGoalsString[1].split(',');
      }
      
      var resultArray = [];  
      for (var i = 0; i < headersArray.length; i++) {
        if(statsArray[i] === '' || statsArray[i] === undefined || statsArray[i] === '--'){
          statsArray[i] = 0;
        }
        if(headersArray[i].search(/Conversions_\d+_LSC/) != -1) {
          var goalID = headersArray[i].replace(/Conversions_(\d+)_LSC/, '$1');
          headersArray[i] = goalsNames[goalsIDs.indexOf(goalID)];          
        }
        resultArray[i] = headersArray[i] + ': ' + statsArray[i]; //Cобираем строку вида "Заголовок: Значение"
      }
      
      var budgetString = 'Amount: ' + budgets[j]; //"Заголовок: Значение" для общего счета
      resultArray.push(budgetString);
      
      returnArray[j] = [lastReportDate + resultArray.join(delimiter).replace(/(\d)\.(\d)/g, '$1,$2')]; //Сбор финальной строки
    }
  }
  return returnArray;
}

//Группируем логины по агентствам
function sortLoginsForBudgets(loginsArray, flagsArray) {
  var loginsSorted = [[],[],[]];
  
  for(var i = 0; i < loginsArray.length; i++) {
    if(flagsArray[i][0] === ''){ //берем только строки без флагов
      if(loginsArray[i][0].match(/ru/) != null) {
        loginsSorted[0].push(loginsArray[i][0]);
      }
      else if(loginsArray[i][0].match(/by/) != null) {
        loginsSorted[1].push(loginsArray[i][0]);          
      }
      else if(loginsArray[i][0].match(/kz/i) != null) {
        loginsSorted[2].push(loginsArray[i][0]);  
      }
    }
  }
  
  for(var a = 0; a < loginsSorted.length; a++) { //Разделяем массивы, если их длина больше 50
    if(loginsSorted[a].length > 50) {
      var removedLogins = loginsSorted[a].splice(50, loginsSorted[a].length - 50); 
      loginsSorted.push(removedLogins);
    }
  }
  
  loginsSorted = loginsSorted.filter(function(element) { //Избавляемся от пустых массивов (когда нет клиентов из определенного агентства)
   return element.length >= 1;
});
  return loginsSorted;
}

//Формируем запросы к API для агентств
function setBudgetsRequests(loginsSorted) {
  var requests = [];
  for(var i = 0; i < loginsSorted.length; i++){
    var data = {
      'method': 'AccountManagement',
      'token': chooseToken(loginsSorted[i][0]), //выбираем токен по первому логину в массиве
      'locale': 'ru',
      'param': {
        'Action': 'Get',
        'SelectionCriteria': {
          'Logins': loginsSorted[i], //все логины агентства
        }
      }       
    };
    requests[i] = {
      'url': 'https://api.direct.yandex.ru/live/v4/json/', //URL  сервиса Яндекс API 
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(data)
    };
  }
  return requests;
}
//Отправляем запросы и получаем данные 
function getBudgets(requests){
  var report = [];
  var response = UrlFetchApp.fetchAll(requests); //отправляем запросы сразу ко всем агентствам
  for(var i = 0; i < response.length; i++){
    report[i] = JSON.parse(response[i].getContentText()).data['Accounts']; //парсим нужную часть ответов в объекты
  }
  for(var j = 1; j < report.length; j++){
    report[0] = report[0].concat(report[j]); //склеиваем массивы объектов в один массив
  }
  return report[0];
}
//Распределяем данные общего счета по нужным строкам
function indexBudgets(loginsArray, budgets){
  var budgetsIndexed = [];
  for(var i = 0; i < loginsArray.length; i++) {
    budgetsIndexed[i] = 0; //Предзаполняем 0-ым значением для строк с флагом
    for(var j = 0; j < budgets.length; j++) {
      if(loginsArray[i][0] == budgets[j]['Login']) //Если логины совпадают, то
        budgetsIndexed[i] = budgets[j]['Amount']; //Присваиваем этому индексу значение поля остатка бюджета
    }
  }
  return budgetsIndexed;
}
//Формируем запросу для отчета по кликам и расходу для каждого логина
function setStatsRequests(loginsArray, startDateArray, goalsCol){
  var todayDate = new Date();
  var requests = [];  
  var delimiter = '/|/';
  for(var i = 0; i < loginsArray.length; i++){ //1 логин - 1 запрос
    var payload = {
    'params': {
      'SelectionCriteria': {
        'DateFrom': Utilities.formatDate(new Date(startDateArray[i][0]), 'GMT+3', 'yyyy-MM-dd'),  // начало периода
        'DateTo': Utilities.formatDate(todayDate, 'GMT+3', 'yyyy-MM-dd') // конец периода (сегодня)
      }, 
      'FieldNames': ['Clicks', 'Cost'], 
      'ReportName': loginsArray[i][0] + '_report_final_' + Utilities.formatDate(todayDate, 'GMT+2', 'dd-MM-yyyy HH:mm'), //формируем уникальное имя отчета с привязкой к логину в времени формирования отчета
      'ReportType': 'ACCOUNT_PERFORMANCE_REPORT', //имя отчета, другие работают сильно медленнее
      'DateRangeType': 'CUSTOM_DATE',      
      'Format': 'TSV',
      'IncludeVAT': 'NO',
      'IncludeDiscount': 'NO'
     }
    }    
    if(!(goalsCol[i][0] === 'Ошибка' || goalsCol[i][0] === '')){
      payload.params['Goals'] = goalsCol[i][0].split(delimiter)[0].split(',');
      payload.params['FieldNames'] = ['Clicks', 'Cost', 'Conversions'];
    }
    
    requests[i] = {
    'url': 'https://api.direct.yandex.com/json/v5/reports', //сервис Яндекс API для получения отчета
    'headers': {
      'Authorization': 'Bearer ' + chooseToken(loginsArray[i][0]), //Выбираем токен по логину
      'Accept-Language': 'ru',
      'Client-Login': loginsArray[i][0], //наш логин
      'returnMoneyInMicros': 'false',
      'skipReportHeader': 'true',
      'skipColumnHeader': 'false',
      'skipReportSummary': 'true',
      'processingMode': 'auto'
      },
    'muteHttpExceptions': true, 
    'payload': JSON.stringify(payload)    
    }
  }
  return requests;
}

function getStats(requests) {  
  var report = [];
  var i = 0;
  var tryCount = 0; //сколько раз нас просили подождать
  var response = UrlFetchApp.fetchAll(requests); //отправляем запросы сразу ко всем логинам
    while(i < requests.length) { 
      if(response[i].getResponseCode() == 200){
        Logger.log('Строка ' + (i+2) + ', отчет готов');
        report[i] = response[i].getAllHeaders()['Date'] + '\n' + response[i].getContentText(); //отчет готов, забираем данные + дату их получения
        i++; //переходим к следующему отчету
        continue;
      } else if (response[i].getResponseCode() == 201) { 
        Logger.log('Строка ' + (i+2) + ', отчет в очереди, ждем: ' + parseInt(response[i].getAllHeaders()['retryin']) + ' сек');
        Utilities.sleep(parseInt(response[i].getAllHeaders()['retryin'])*1000); //отчет находится в очереди на формирование, ждем рекомендуемое количество секунд
      } else if (response[i].getResponseCode() == 202 && tryCount <= 6) { 
        Logger.log('Строка ' + (i+2) + ', отчет в обработке, ждем: ' + parseInt(response[i].getAllHeaders()['retryin']) + ' сек');
        Utilities.sleep(parseInt(response[i].getAllHeaders()['retryin'])*1000); //отчет в процессе формирования, ждем рекомендуемое количество секунд
        tryCount++; 
      } else if (response[i].getResponseCode() == 202 && tryCount > 6) { //мы ждали слишком много времени, пропускаем все отчеты, что не успели сформироваться
        Logger.log('Строка ' + (i+2) + ', отчет не готов - пропускаем'); 
        report[i] = 'Skipped'; 
        i++;
        continue;
      } else {
        Logger.log('Ошибка ' + response[i]);
        report[i] = 'Skipped';
        i++;
      }             
      response = UrlFetchApp.fetchAll(requests); //мы подождали и снова проверяем статусы отчетов
  }
  return report;
}

function checkCounterLinks(countersCol, loginsCol, sheet){
  for(var i = 0; i < countersCol.length;  i++){
    if(countersCol[i][0] === ''){
      Logger.log('Формирую ссылку на Метрику для ' + loginsCol[i][0]);
      var counterLink = 'metrika.yandex.ru/stat/sources?group=day&period=week&attribution=LastSign&id=' + getCounterID(loginsCol[i][0]);
      sheet.getRange('H' + (2 + i)).setValue(counterLink);
    }
  }
}

function checkGoals(goalsCol, loginsCol, sheet){
  for(var i = 0; i < goalsCol.length;  i++){
    if(goalsCol[i][0] === ''){
      Logger.log('Получаю цели для ' + loginsCol[i][0]);
      goalsCol[i][0] = getGoals(loginsCol[i][0]);
      sheet.getRange('A' + (2 + i)).setValue(goalsCol[i][0]);
    }
  }
  return goalsCol;
}

function getGoals(login){
  
  var counterID = getCounterID(login);
  
  var goalsURL = 'https://api-metrika.yandex.net/management/v1/counter/' + counterID + '/goals?useDeleted=false';
  var request = {
    'headers': {
      'Authorization': 'OAuth ' + chooseToken(login),      
    },
    'muteHttpExceptions': true,
  }
  var response = UrlFetchApp.fetch(goalsURL, request);
  var goalsIDs, goalsNames, result;
  var delimiter = '/|/';
  if(response.getResponseCode() == 200){
    var report = response.getContentText();
    
    goalsIDs = report.match(/"id":\d+/g).join(',').replace(/"id":/g, ''); 
    goalsNames = report.match(/"name":".+?"/g).join(',').replace(/"|name|:|\s\((a|а)\)/g, '');
    result = goalsIDs + delimiter + goalsNames;
  } else {
        Logger.log('failed ' + response);
        result = 'Ошибка';
      }             
  return result;
}

function getCounterID(login){
  var campaignsURL = 'https://api.direct.yandex.com/json/v5/campaigns';  
  
  var payload = {
    'method': 'get',
    'params': {
      'SelectionCriteria': {},
      'FieldNames': ['Id'],
      'TextCampaignFieldNames': ['CounterIds']
    }
  }    
  
  var request = {
    'headers': {
      'Authorization': 'Bearer ' + chooseToken(login), //Выбираем токен по логину
      'Accept-Language': 'ru',
      'Client-Login': login, //наш логин
    },
    'muteHttpExceptions': true, 
    'payload': JSON.stringify(payload)
  }
  
  var response = UrlFetchApp.fetch(campaignsURL, request);
  var counterID = (response.getContentText()).match(/{"Items":\[(\d+)/)[1];
  return counterID;
}

function checkClientName(clientNameCol, loginsCol, sheet) {
  for(var i = 0; i < clientNameCol.length;  i++){
    if(clientNameCol[i][0] === ''){
      Logger.log('Запрашиваю имя клиента для логина ' + loginsCol[i][0]);
      var clientName = getClientName(loginsCol[i][0]);
      sheet.getRange('D' + (2 + i)).setValue(clientName);
    }
  }
}

function getClientName(login) {
   var clientsURL = 'https://api.direct.yandex.com/json/v5/clients';  
  
  var payload = {
    'method': 'get',
    'params': {
      'FieldNames': ['ClientInfo'],
    }
  }    
  
  var request = {
    'headers': {
      'Authorization': 'Bearer ' + chooseToken(login), //Выбираем токен по логину
      'Accept-Language': 'ru',
      'Client-Login': login, //наш логин
    },
    'muteHttpExceptions': true, 
    'payload': JSON.stringify(payload)
  }
  
  var response = UrlFetchApp.fetch(clientsURL, request);
  Logger.log(response.getContentText());
  var clientName = (response.getContentText()).match(/ClientInfo":"(.+?)"}/)[1].replace(/\\"/g, "\"");
  return clientName;
}















