function contractPaymentSchedule() {
  paymentSchedule("contract")
}

function quotationPaymentSchedule(){
  paymentSchedule("quotation")
}

function paymentSchedule(corq){
  var sheet = SpreadsheetApp.openById(FUF.contractSsID[env]).getSheetByName(corqGenerateSheetName[corq]);

  var menus = FUF.removeEmptyRows(sheet.getRange(serviceInfoCoordinate[corq]).getValues().map((element) => FUF.removeBlankElements(element)));
  var paymentOption = sheet.getRange(paymentOptionRange[corq]).getValues()[0][0];

  console.log(paymentOption)

  // var serviceStartDates = FUF.removeBlankElements(FUF.transpose(sheet.getRange(serviceStartDateCoordinate[corq]).getValues())[0]);
  // var serviceStartDate = serviceStartDates.reduce(function (min, currentDate) {
  //   return (min < currentDate) ? min : currentDate;
  // });;

  var serviceEndDates = FUF.removeBlankElements(FUF.transpose(sheet.getRange(serviceEndDateCoordinate[corq]).getValues())[0]);
  var serviceEndDate = serviceEndDates.reduce(function (max, currentDate) {
    return (max > currentDate) ? max : currentDate;
  });;


  var contractDate = sheet.getRange(contractDateCoordinate).getValues();

  var firstDate = new Date(contractDate[0][0].getTime())
  firstDate.setDate(firstDate.getDate() + 3);
  paymentScheduleDueDates = [firstDate]
  // console.log(contractDate[0][0])

  var lastDateOfFirstMonth = new Date(firstDate);;
  lastDateOfFirstMonth.setMonth(lastDateOfFirstMonth.getMonth() + 1);
  lastDateOfFirstMonth.setDate(0)

  var lastPaymentDueDate = new Date(serviceEndDate.getTime());
  lastPaymentDueDate.setDate(monthlyPaymentDueDate)
  var latestDueDate = new Date(firstDate);

  if (paymentOption == "月払い") {
    lastPaymentDueDate.setMonth(lastPaymentDueDate.getMonth() - 1)
    latestDueDate.setMonth(latestDueDate.getMonth() - 1);
    for (var i = 0; latestDueDate < lastPaymentDueDate; i++) {
      latestDueDate.setMonth(latestDueDate.getMonth() + 1);
      latestDueDate.setDate(monthlyPaymentDueDate)
      if (latestDueDate > lastDateOfFirstMonth) { //
        paymentScheduleDueDates.push(new Date(latestDueDate))
      }
    }
  } else if (paymentOption == "６ヶ月ごと分割") {
    lastPaymentDueDate.setMonth(lastPaymentDueDate.getMonth() - 6)
    latestDueDate.setMonth(latestDueDate.getMonth() - 1);
    for (var i = 0; latestDueDate < lastPaymentDueDate; i++) {
      latestDueDate.setMonth(latestDueDate.getMonth() + 6);
      latestDueDate.setDate(monthlyPaymentDueDate)
      if (latestDueDate > lastDateOfFirstMonth) {
        paymentScheduleDueDates.push(new Date(latestDueDate))
      }
    }
  }
  else if (paymentOption == "３ヶ月ごと分割") {
    lastPaymentDueDate.setMonth(lastPaymentDueDate.getMonth() - 3)
    latestDueDate.setMonth(latestDueDate.getMonth() - 1);
    for (var i = 0; latestDueDate < lastPaymentDueDate; i++) {
      latestDueDate.setMonth(latestDueDate.getMonth() + 3);
      latestDueDate.setDate(monthlyPaymentDueDate)
      if (latestDueDate > lastDateOfFirstMonth) {
        paymentScheduleDueDates.push(new Date(latestDueDate))
      }
    }
  }

  // console.log(serviceEndDate)
  // console.log(lastPaymentDueDate)
  // console.log(paymentScheduleDueDates)
  // console.log(menus)

  var paymentSchedule = paymentCalculateBySchedule(menus, paymentScheduleDueDates)
  // console.log(paymentSchedule)

  FUF.setValuesWithPadding(sheet.getRange(paymentScheduleDatesCoordinate[corq]),[paymentSchedule[0]] )

  FUF.setValuesWithPadding(sheet.getRange(paymentScheduleAmountCoordinate[corq]),paymentSchedule.slice(1));
}

function paymentCalculateBySchedule(menus, paymentScheduleDueDates) {
  var paymentSchedule = [paymentScheduleDueDates];
  // サービス項目ごとの繰り返し
  for (var i = 0; i < menus.length; i++) {
    monthlyAmounts = calculateMonthlyAmounts(menus[i][6], menus[i][2], menus[i][3])
    var paymentParServiceAndDueDate = []
    // 支払い期限日ごとの繰り返し
    for (var j = 0; j < paymentScheduleDueDates.length - 1; j++) {
      var paymentParDueDate = 0
      // サービス利用月ごとの繰り返し
      for (var k = 0; k < monthlyAmounts[1].length; k++) {
        if (monthlyAmounts[0][k] < paymentScheduleDueDates[j + 1]) {
          paymentParDueDate = paymentParDueDate + monthlyAmounts[1][k]
          monthlyAmounts[1][k] = 0;
        }
      }
      paymentParServiceAndDueDate.push(paymentParDueDate);
    }
    paymentParServiceAndDueDate.push(monthlyAmounts[1].reduce(function (sum, number) {
      return sum + number;
    }, 0));

    paymentSchedule.push(paymentParServiceAndDueDate.map(function(element) {return element * 1.10}))
  }
  return paymentSchedule
}

