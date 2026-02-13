// static/auto_report.js
(() => {
  // Чтобы не запускалось много раз при навигации внутри SPA
  const RUN_KEY = "miac_auto_report_running_v1";
  if (window.sessionStorage.getItem(RUN_KEY) === "1") {
    return;
  }
  window.sessionStorage.setItem(RUN_KEY, "1");

  const BODYS = [
    ['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"ч","expression":"[Kolichestvo_chernovikov]","sort":0,"aggregator":1}],"id":"e7a56ae1d48c4fd9816ee80f7ed8f1cf","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"8608ab7565b645f98fe9ad899797acbe","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"8608ab7565b645f98fe9ad899797acbe"}', "ПОПЫТОК ЗАПИСИ, прерванных пользователями"],
    ['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"ош","expression":"[Kolichestvo_oshibok]","sort":0,"aggregator":1}],"id":"02b8f81f317546dd876a84c961a7371e","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"7126892d8fcf422296f79d1d6f668df2","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"7126892d8fcf422296f79d1d6f668df2"}', "НЕУДАЧНЫХ попыток записи"],
    ['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"org","expression":"[Dostupnoe_raspisanie_ne_naideno]+[Dostupnie_dolzhnosti_ne_naideni]+[Dostupnie_medspetsialisti_ne_nai]+[Dostupnie_MO_ne_naideni]+[Dostupnie_napravleniya_ne_naiden]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"cabe0dac6c0c41ecac6846d8c01de54c","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"cabe0dac6c0c41ecac6846d8c01de54c"}', "НЕУДАЧНЫХ ПОПЫТОК ЗАПИСАТЬСЯ по организационным причинам в МО"],
    ['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"teh","expression":"[Vnutrennyaya_oshibka_RMIS]+[Otsutstvie_otveta_ot_RMIS]+[Oshibka_validatsii_shemi]+[Servis_RMIS_priostanovlen]+[Taimaut_sessii]+[Patsient_ne_naiden_v_RMIS]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"89d053b0d6a64677982a2754c92631c3","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"89d053b0d6a64677982a2754c92631c3"}', "НЕУДАЧНЫХ ПОПЫТОК ЗАПИСАТЬСЯ по техническим причинам"],
    ['{"QueryType":"GetOlapData+Query","OlapSettings": {"olapGroups": [{"measures": [{"id": "Kolichestvo_zapisei","sort": 0,"aggregator": 0}],"attributeGroupsInRows": [],"attributeGroupsInColumns": [],"filters": [[{"selectedFilterValues": {"Values": [["Челябинская область"]],"Range": {"Min": null,"Max": null},"UseRange": false,"UseExcluding": false},"attribute": {"id": "Subekt_RF","dimensionOrDimensionRoleId": {"type": "DimensionIdDto","kind": "DimensionIdDto","value": "db3_new_fer_quality"}}},{"selectedFilterValues": {"Values": [["ORG"]],"Range": {"Min": null,"Max": null},"UseRange": false,"UseExcluding": false},"attribute": {"id": "sp","dimensionOrDimensionRoleId": {"type": "DimensionIdDto","kind": "DimensionIdDto","value": "db3_new_fer_quality"}}}]],"dateRangeFilters": [{"start": {"type": 0,"fixedDateRange": {"date": "DSTARTT00:00:00.000Z"},"relativeDateRange": {"offsetPeriod": 0,"dateStartingPointType": 1,"relativeDateRangeType": 1,"offset": 0}},"end": {"type": 0,"fixedDateRange": {"date": "DSTOPT00:00:00.000Z"},"relativeDateRange": {"offsetPeriod": 0,"dateStartingPointType": 1,"relativeDateRangeType": 1,"offset": 0}},"dimensionOrDimensionRoleId": {"type": "DimensionRoleIdDto","kind": "DimensionRoleIdDto","value": "Data"}}],"calculatedMeasures": [],"id": "051c09be1301484cb053222f51b3370e","measureGroupId": "db3_sp_znp_fer_quality","rowTotal": false,"columnTotal": false,"showAllDimensionValues": false}],"databaseId": "Indeks","limit": {"rowsLimit": 10000,"columnsLimit": 50}},"CalculationQueries": [],"AdditionalLogs": {"widgetGuid": "3bbe5d7cac944bbe8f8593ea347ee4cc","dashboardGuid": "8d82093225eb470595ae4d49d2edc555","sheetGuid": "75e7b48008b34db482c350b2333e2d45"}, "WidgetGuid": "3bbe5d7cac944bbe8f8593ea347ee4cc"}', "УСПЕШНЫХ записей"],
    ['{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[],"attributeGroupsInRows":[],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["ORG"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"DSTARTT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"DSTOPT00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[{"name":"fed","expression":"[Vnutrennyaya_oshibka_KU_FER]+[Oshibka_potrebitelya]+[Taimaut_sessii]","sort":0,"aggregator":1}],"id":"45a0041e292649ec80160faabc34cfbd","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":10000,"columnsLimit":50}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"8aa36e8086104c779fd5cad55dcb73cc","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"8aa36e8086104c779fd5cad55dcb73cc"}', "ошибки федерального уровня"]
  ];

  // ВАЖНО: url делаем относительным, чтобы шло через ваш прокси /bi/...
  const url = "/bi/corelogic/api/query";

  // _ORG_ оставляем как в расширении (только даты можно менять как хотите)
  const _ORG_ = '{"QueryType":"GetOlapData+Query","OlapSettings":{"olapGroups":[{"measures":[{"id":"Kolichestvo_zapisei","sort":0,"aggregator":0}],"attributeGroupsInRows":[{"attributes":[{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}],"total":false}],"attributeGroupsInColumns":[],"filters":[[{"selectedFilterValues":{"Values":[["Челябинская область"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":false},"attribute":{"id":"Subekt_RF","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}},{"selectedFilterValues":{"Values":[["<Пусто>"]],"Range":{"Min":null,"Max":null},"UseRange":false,"UseExcluding":true},"attribute":{"id":"sp","dimensionOrDimensionRoleId":{"type":"DimensionIdDto","kind":"DimensionIdDto","value":"db3_new_fer_quality"}}}]],"dateRangeFilters":[{"start":{"type":0,"fixedDateRange":{"date":"2025-05-19T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"end":{"type":0,"fixedDateRange":{"date":"2025-05-25T00:00:00.000Z"},"relativeDateRange":{"offsetPeriod":0,"dateStartingPointType":1,"relativeDateRangeType":1,"offset":0}},"dimensionOrDimensionRoleId":{"type":"DimensionRoleIdDto","kind":"DimensionRoleIdDto","value":"Data"}}],"calculatedMeasures":[],"id":"9b25ea63865d43cd9ee397490bb541b5","measureGroupId":"db3_sp_znp_fer_quality","rowTotal":false,"columnTotal":false,"showAllDimensionValues":false}],"databaseId":"Indeks","limit":{"rowsLimit":300,"columnsLimit":0}},"CalculationQueries":[],"AdditionalLogs":{"widgetGuid":"43c359a9b3f847c8bc309044615168e4","dashboardGuid":"8d82093225eb470595ae4d49d2edc555","sheetGuid":"75e7b48008b34db482c350b2333e2d45"},"WidgetGuid":"43c359a9b3f847c8bc309044615168e4"}';

  function replaceAll(str, search, repl) {
    return str.split(search).join(repl);
  }

  function downloadFile(name, value) {
    const blob = new Blob([value ?? ""], { type: "application/json;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      document.body.removeChild(a);
    }, 2000);
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function getAccessToken() {
    try {
      const raw = sessionStorage.getItem("oidc.user:/idsrv:DashboardsApp");
      if (!raw) return null;
      const obj = JSON.parse(raw);
      return obj?.access_token ?? null;
    } catch {
      return null;
    }
  }

  async function waitFor(conditionFn, { timeoutMs = 180000, stepMs = 500 } = {}) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      try {
        const v = await conditionFn();
        if (v) return v;
      } catch (_) {}
      await delay(stepMs);
    }
    return null;
  }

  function getIframeDoc() {
    const iframe = document.getElementsByTagName("iframe")[0];
    if (!iframe) return null;
    try {
      return iframe.contentWindow?.document ?? null;
    } catch {
      return null;
    }
  }

  function readDateRangeFromIframe(iframeDoc) {
    const el = iframeDoc.getElementsByClassName("datepicker-here va-date-filter")[0];
    if (!el || !el.value) return null;

    // формат как в расширении: "dd.mm.yyyy - dd.mm.yyyy"
    const parts = el.value.split("-");
    if (parts.length < 2) return null;

    const dStart = parts[0].trim().split(".");
    const dStop = parts[1].trim().split(".");

    const startIso = `${dStart[2]}-${dStart[1]}-${dStart[0]}`;
    const stopIso = `${dStop[2]}-${dStop[1]}-${dStop[0]}`;
    return { startIso, stopIso };
  }

  async function fetchJson(body, token) {
    const resp = await fetch(url, {
      method: "POST",
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      body
    });
    return await resp.json();
  }

  async function getReq(org, dStart, dStop, token) {
    const req = [];
    for (let j = 0; j < BODYS.length; j++) {
      let b = replaceAll(BODYS[j][0], "ORG", org);
      b = replaceAll(b, "DSTART", dStart);
      b = replaceAll(b, "DSTOP", dStop);
      req.push(fetchJson(b, token));
      await delay(50);
    }
    return req;
  }

  async function run() {
    // 1) ждём токен (после авторизации)
    const token = await waitFor(() => getAccessToken(), { timeoutMs: 180000, stepMs: 500 });
    if (!token) {
      console.warn("[MIAC] token not found");
      window.sessionStorage.removeItem(RUN_KEY);
      return;
    }

    // 2) ждём iframe и элементы внутри
    const iframeDoc = await waitFor(() => {
      const d = getIframeDoc();
      if (!d) return null;
      // проверим, что хотя бы datepicker доступен
      const ok = d.getElementsByClassName("datepicker-here va-date-filter")[0];
      return ok ? d : null;
    }, { timeoutMs: 180000, stepMs: 500 });

    if (!iframeDoc) {
      console.warn("[MIAC] iframe doc not ready");
      window.sessionStorage.removeItem(RUN_KEY);
      return;
    }

    // 3) диапазон дат (как в расширении — читаем с виджета)
    const range = readDateRangeFromIframe(iframeDoc);
    if (!range) {
      console.warn("[MIAC] date range not found");
      window.sessionStorage.removeItem(RUN_KEY);
      return;
    }

    const { startIso: dStart, stopIso: dStop } = range;

    // 4) получаем список организаций тем же способом, что и расширение
    const orgResp = await fetchJson(_ORG_, token);
    const org = [];
    const rows = orgResp?.result?.dataFrame?.rows ?? [];
    for (let i = 0; i < rows.length; i++) {
      org.push(rows[i][0]);
    }

    const total = org.length;
    if (!total) {
      alert("Организации не найдены");
      window.sessionStorage.removeItem(RUN_KEY);
      return;
    }

    // где показывать прогресс (оставлено как в расширении)
    function setProgress(text) {
      try {
        const el = iframeDoc.getElementsByClassName("va-widget-body-container")[21]
          ?.getElementsByTagName("div")[1];
        if (el) el.innerText = text;
      } catch (_) {}
    }

    const result = [];
    const loop = 10;

    for (let i = 0; i < total; i++) {
      setProgress(`ПРОГРЕСС: ${i + 1}/${total}`);

      const orgName = org[i];
      let safeOrg = replaceAll(orgName, '"', '\\"');

      let response1 = [orgName];

      // Повторяем логику ретраев из расширения
      let resPromises = await getReq(safeOrg, dStart, dStop, token);

      for (let x = 0; x < loop; x++) {
        let ok = true;
        const r = await Promise.all(resPromises);

        for (let w = 0; w < r.length; w++) {
          if (r[w]?.hasError) {
            ok = false;
            if (loop - 1 === x) {
              response1.push([BODYS[w][1], r[w]?.errorMessage ?? "Ошибка"]);
            } else {
              // перезапуск для этой организации
              response1 = [orgName];
              resPromises = await getReq(safeOrg, dStart, dStop, token);
              break;
            }
          } else {
            response1.push([BODYS[w][1], r[w]?.result?.dataFrame?.values?.[0]]);
          }
        }

        if (ok) break;
      }

      result.push(response1);

      if (i === total - 1) {
        downloadFile("parse.json", JSON.stringify(result));
        alert("Работа завершена");
        window.sessionStorage.removeItem(RUN_KEY);
      }
    }
  }

  // Запуск: после полной загрузки страницы и небольшого ожидания
  window.addEventListener("load", () => {
    setTimeout(() => {
      run().catch((e) => {
        console.error("[MIAC] run failed", e);
        window.sessionStorage.removeItem(RUN_KEY);
      });
    }, 1500);
  });
})();
