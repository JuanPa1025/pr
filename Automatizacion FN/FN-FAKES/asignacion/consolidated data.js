function actualizeQuantityByMonth() {
  const lastMonth = new Date().getMonth() - 1; // Last month

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analystSheet = ss.getSheetByName("analistas");
  const consolidateSheet = ss.getSheetByName("consolidado");

  const consolidateData = consolidateSheet.getRange(2, 1, consolidateSheet.getLastRow() - 1, 9).getValues();

  const uniqueUsers = [];
  for (const row of consolidateData) {
    const user = row[5];
    if (!uniqueUsers.includes(user)) {
      uniqueUsers.push(user);
    }
  }

  const quantityByUser = [];
  for (const pyUser of uniqueUsers) {
    const filterByUser = consolidateData.filter(data => data[5] === pyUser && (parseInt(data[0].split('/')[1])) === lastMonth + 1);
    const workedDays = [... new Set(filterByUser.map(data => data[0]))].length;
    const monthName = Utilities.formatDate(new Date(new Date().setMonth(lastMonth)), "GMT-3", "MMMM");
    const userData = {
      user: pyUser,
      month: monthName,
      average: 0,
      total: 0,
      quantity: filterByUser.length,
      notNegativequantity: 0,
      workedDays,
    };
    for (const row of filterByUser) {
      const duration = row[8];
      const [hh, mm, ss] = duration.split(':');
      const seconds = parseInt(ss) + (parseInt(mm) * 60) + (parseInt(hh) * 60 * 60);

      if (seconds > 0) {
        userData.total = userData.total + seconds;
        userData.notNegativequantity++;
      }
    }
    if (userData.quantity > 0) {
      quantityByUser.push(userData);
    }
  }
  const metricsData = quantityByUser.map(userData => ({
    ...userData,
    totalRaw: userData.total,
    averageRaw: userData.notNegativequantity > 0 ? Math.floor(userData.total / userData.notNegativequantity) : 0,
    average: userData.notNegativequantity > 0 ? secondsToHourFormat(Math.floor(userData.total / userData.notNegativequantity)) : "0:00:00",
    total: userData.total > 0 ? secondsToHourFormat(userData.total) : "0:00:00",
  }));

  console.log({ metricsData });

  const placeFormatData = metricsData.map(data => [data.user, data.month, data.average, data.total, data.quantity, data.workedDays])

  if (placeFormatData.length) {
    analystSheet.getRange(analystSheet.getLastRow() + 1, 1, placeFormatData.length, placeFormatData[0].length).setValues(placeFormatData);
  }
}

function secondsToHourFormat(seconds) {
  const sec = Number(seconds);
  const hh = Math.floor(sec / 3600);
  const mm = Math.floor(sec % 3600 / 60);
  const ss = Math.floor(sec % 3600 % 60);

  return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`
}