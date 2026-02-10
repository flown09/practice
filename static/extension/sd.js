const BODYS = [
['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"ч","expression":"[Kolichestvo_chernovikov]","sort":0,"aggregator":1}],"id":"e7a56ae1d48c4fd9816ee80f7ed8f1cf","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"8608ab7565b645f98fe9ad899797acbe","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"8608ab7565b645f98fe9ad899797acbe"}', 'ПОПЫТОК ЗАПИСИ, прерванных пользователями'
],
['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"ош","expression":"[Kolichestvo_oshibok]","sort":0,"aggregator":1}],"id":"02b8f81f317546dd876a84c961a7371e","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"7126892d8fcf422296f79d1d6f668df2","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"7126892d8fcf422296f79d1d6f668df2"}', 'НЕУДАЧНЫХ попыток записи'
],
['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"org","expression":"[Dostupnoe_raspisanie_ne_naideno]+[Dostupnie_dolzhnosti_ne_naideni]+[Dostupnie_medspetsialisti_ne_nai]+[Dostupnie_MO_ne_naideni]+[Dostupnie_napravleniya_ne_naiden]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"cabe0dac6c0c41ecac6846d8c01de54c","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"cabe0dac6c0c41ecac6846d8c01de54c"}', 'НЕУДАЧНЫХ ПОПЫТОК ЗАПИСАТЬСЯ по организационным причинам в МО'
],
['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"teh","expression":"[Vnutrennyaya_oshibka_RMIS]+[Otsutstvie_otveta_ot_RMIS]+[Oshibka_validatsii_shemi]+[Servis_RMIS_priostanovlen]+[Taimaut_sessii]+[Patsient_ne_naiden_v_RMIS]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"89d053b0d6a64677982a2754c92631c3","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"89d053b0d6a64677982a2754c92631c3"}', 'НЕУДАЧНЫХ ПОПЫТОК ЗАПИСАТЬСЯ по техническим причинам'
],
['{"QueryType":"GetOlapData+Query","OlapSettings": {"olapGroups": [{"measures": [{"id": "Kolichestvo_zapisei","sort": 0,"aggregator": 0}],"attributeGroupsInRows": [],"attributeGroupsInColumns": [],"filters": [[{"selectedFilterValues": {"Values": [["Челябинская область"]],"Range": {"Min": null,"Max": null},"UseRange": false,"UseExcluding": false},"attribute": {"id": "Subekt_RF","dimensionOrDimensionRoleId": {"type": "DimensionIdDto","kind": "DimensionIdDto","value": "db3_new_fer_quality"}}},{"selectedFilterValues": {"Values": [["ORG"]],"Range": {"Min": null,"Max": null},"UseRange": false,"UseExcluding": false},"attribute": {"id": "sp","dimensionOrDimensionRoleId": {"type": "DimensionIdDto","kind": "DimensionIdDto","value": "db3_new_fer_quality"}}}]],"dateRangeFilters": [{"start": {"type": 0,"fixedDateRange": {"date": "DSTARTT00:00:00.000Z"},"relativeDateRange": {"offsetPeriod": 0,"dateStartingPointType": 1,"relativeDateRangeType": 1,"offset": 0}},"end": {"type": 0,"fixedDateRange": {"date": "DSTOPT00:00:00.000Z"},"relativeDateRange": {"offsetPeriod": 0,"dateStartingPointType": 1,"relativeDateRangeType": 1,"offset": 0}},"dimensionOrDimensionRoleId": {"type": "DimensionRoleIdDto","kind": "DimensionRoleIdDto","value": "Data"}}],"calculatedMeasures": [],"id": "051c09be1301484cb053222f51b3370e","measureGroupId": "db3_sp_znp_fer_quality","rowTotal": false,"columnTotal": false,"showAllDimensionValues": false}],"databaseId": "Indeks","limit": {"rowsLimit": 10000,"columnsLimit": 50}},"CalculationQueries": [],"AdditionalLogs": {"widgetGuid": "3bbe5d7cac944bbe8f8593ea347ee4cc","dashboardGuid": "8d82093225eb470595ae4d49d2edc555","sheetGuid": "75e7b48008b34db482c350b2333e2d45"}, "WidgetGuid": "3bbe5d7cac944bbe8f8593ea347ee4cc"}', 'УСПЕШНЫХ записей'
],
['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"2023-09-01T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"2023-09-06T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"fed","expression":"[Vnutrennyaya_oshibka_KU_FER]+[Oshibka_potrebitelya]+[Taimaut_sessii]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"8aa36e8086104c779fd5cad55dcb73cc","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"8aa36e8086104c779fd5cad55dcb73cc"}','ошибки федерального уровня'
]
];

//var _ORG_ = '{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[{"id":"Kolichestvo_zapisei","sort":0,"aggregator":0}],"attributeGroupsInRows":[{"attributes":[{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}],"total":false}],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["<Пусто>"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":true},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"2023-03-10T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"2023-03-16T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[],"id":"9b25ea63865d43cd9ee397490bb541b5","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":300,"columnsLimit":0}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"43c359a9b3f847c8bc309044615168e4","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"43c359a9b3f847c8bc309044615168e4"}';
//var _ORG_ = '{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[{"id":"Kolichestvo_zapisei","sort":0,"aggregator":0}],"attributeGroupsInRows":[{"attributes":[{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}],"total":false}],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["<Пусто>"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":true},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"2023-12-04T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"2023-12-10T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[],"id":"9b25ea63865d43cd9ee397490bb541b5","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":300,"columnsLimit":0}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"43c359a9b3f847c8bc309044615168e4","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"43c359a9b3f847c8bc309044615168e4"}';
var _ORG_ = '{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[{"id":"Kolichestvo_zapisei","sort":0,"aggregator":0}],"attributeGroupsInRows":[{"attributes":[{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}],"total":false}],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["<Пусто>"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":true},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"2025-05-19T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"2025-05-25T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[],"id":"9b25ea63865d43cd9ee397490bb541b5","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":300,"columnsLimit":0}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"43c359a9b3f847c8bc309044615168e4","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"43c359a9b3f847c8bc309044615168e4"}'
var options = {
	method: 'POST',
	headers: {
	  'Accept': 'application/json',
	  'Content-Type': 'application/json',
	  'Authorization': 'Authorization'
	},
	body: "body",
};

function sleep(milliseconds) {       
    const date = Date.now();        
    let currentDate = null;       
    do {               
       currentDate = Date.now();      
    } while (currentDate - date < milliseconds); 
} 

var url = ["https://info-bi-db.egisz.rosminzdrav.ru/corelogic/api/query"];

function writeFile(name, value) {
  var val = value;

  if (value === undefined) {
    val = "";
  }
  
  var download = document.createElement("a");
  download.href = "data:text/plain;content-disposition=attachment;filename=file," + val;
  download.download = name;
  download.style.display = "none";
  download.id = "download"; document.body.appendChild(download);
  document.getElementById("download").click();
  document.body.removeChild(download);
}

async function getReq(o, dStart, dStop) {
	var token = JSON.parse(sessionStorage.getItem('oidc.user:/idsrv:DashboardsApp'))["access_token"];
	var req = [];

	for (var j = 0; j < BODYS.length; j++) {
		var b = replaceAll(BODYS[j][0], "ORG", o);
		b = replaceAll(b, "DSTART", dStart);
		b = replaceAll(b, "DSTOP", dStop);

		options["body"] = b;
		options["headers"]["Authorization"] = 'Bearer ' + token;

		console.log(options);
		req.push(fetch(url, options).then(response => response.json()));
		sleep(50);
	}
	
	return req;
}

var autoExportStarted = false;

function startAutoExport() {
	if (autoExportStarted) {
		return;
	}

	autoExportStarted = true;

		var token = JSON.parse(sessionStorage.getItem('oidc.user:/idsrv:DashboardsApp'))["access_token"];
		var org = [];
		
		options["body"] = _ORG_;
		options["headers"]["Authorization"] = 'Bearer ' + token;
		
		(async () => {
			var result = await fetch(url, options);
			var o = await result.json();
			
			for (var i = 0; i < o["result"]["dataFrame"]["rows"].length; i++) {
				org.push(o["result"]["dataFrame"]["rows"][i][0]);
			}
		// })();
		
			// var l = document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("rb-filter-list-container")[1].getElementsByClassName("rb-filter-list-item-text").length;
			var l = org.length;
			
			// for (var i = 0; i < l; i++) {
				// org.push(document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("rb-filter-list-container")[1].getElementsByClassName("rb-filter-list-item-text")[i].textContent);
			// }
			
			var loop = 10;
			var result = [];
			
			var dStart = document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("datepicker-here va-date-filter")[0].value.split("-")[0].trim().split(".");
			dStart = dStart[2] + "-" + dStart[1] + "-" + dStart[0];
			var dStop = document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("datepicker-here va-date-filter")[0].value.split("-")[1].trim().split(".");
			dStop = dStop[2] + "-" + dStop[1] + "-" + dStop[0];

			// l = 2;

			for (var i = 0; i < l; i++) {
				document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("va-widget-body-container")[21].getElementsByTagName("div")[1].innerText = "ПРОГРЕСС: " + (i + 1) + "/" + l;

				var response1 = [];
				response1.push(org[i]);
				o = replaceAll(org[i], '"', '\\"');
			    // o = org[i];
			  
			    res = await getReq(o, dStart, dStop);
			  
				for (var x = 0; x < loop; x++) {
					var st = false;
					var r = await Promise.all(res);
					console.log(r);
					
					for (var w = 0; w < r.length; w++) {
						if (r[w]["hasError"]) {
							// debugger;
							
							if (loop - 1 == x) {
								response1.push([BODYS[w][1], r[w]["errorMessage"]]);
								//console.log(response1);
							} else {
								response1 = [];
								response1.push(org[i])
								res = await getReq(o, dStart, dStop);
								break;
							}
						} else {
							response1.push([BODYS[w][1], r[w]["result"]["dataFrame"]["values"][0]]);
							st = true;
							//break;
						}
					}
					
					console.log(response1);
					
					if (st) {
						break;
					}
				}
			  
			result.push(response1);
			console.log(JSON.stringify(result));
			
			if (i == l - 1) {
				writeFile('parse.json',  await JSON.stringify(result));
				//writeFile('date.txt', dStart + " " + dStop)
				alert("Работа завершена");
			}
		  }
		  })();
}

function checkLoad(container) {
	// debugger;
	if (document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("rb-actions-buttons-container")[1] != undefined) {
		if (document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("rb-actions-buttons-container")[1].getElementsByClassName("rb-filter-cancel-button button")[1] == undefined) {			
			document.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByClassName("rb-actions-buttons-container")[1].append(container);
		}
	}	
}

function replaceAll(string, search, replace) {
  return string.split(search).join(replace);
}

function isDashboardReady() {
	try {
		var authRaw = sessionStorage.getItem('oidc.user:/idsrv:DashboardsApp');
		if (!authRaw) {
			return false;
		}

		var auth = JSON.parse(authRaw);
		if (!auth || !auth["access_token"]) {
			return false;
		}

		var iframe = document.getElementsByTagName('iframe')[0];
		if (!iframe || !iframe.contentWindow || !iframe.contentWindow.document) {
			return false;
		}

		return iframe.contentWindow.document.getElementsByClassName("datepicker-here va-date-filter")[0] != undefined;
	} catch (e) {
		return false;
	}
}

function waitForAuthAndDashboard() {
	var attempts = 0;
	var maxAttempts = 300;
	var timer = setInterval(function() {
		attempts += 1;

		if (isDashboardReady()) {
			clearInterval(timer);
			startAutoExport();
			return;
		}

		if (attempts >= maxAttempts) {
			clearInterval(timer);
			console.warn("Автовыгрузка не запущена: дашборд или авторизация не готовы");
		}
	}, 1000);
}

if (document.readyState === "complete") {
	waitForAuthAndDashboard();
} else {
	window.addEventListener("load", waitForAuthAndDashboard);
}
