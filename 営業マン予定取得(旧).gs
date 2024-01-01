// const spreadSheet = SpreadsheetApp.getActiveSpreadsheet(); 
// const sheetSalesmanId = spreadSheet.getSheetByName("IDシート");
// const sheetLog = spreadSheet.getSheetByName("営業マン予定(詳細含む)5/22~");


// // const dates = []; // 日付オブジェクトを格納する配列
// // const now = new Date();
// // const yesterday = new Date(now);
// // yesterday.setDate(now.getDate() - 1);


// // 任意の期間で取得！⇩
// const dates = []; // 日付オブジェクトを格納する配列
// const dayStart = new Date(2023, 5 - 1, 22); // 検索開始時刻
// const dayEnd   = new Date(2023, 6 - 1, 25); // 検索終了時刻




// function getEveryoneEvents(){
//   // dates.push(new Date(yesterday));
//   for (let date = dayStart; date <= dayEnd; date.setDate(date.getDate() + 1)) {
//       dates.push(new Date(date));
//   }

//   //dayループ
//   for (let i = 0; i <= dates.length; i++){
//     const timeStart = new Date(dates[i])
//     Logger.log(timeStart)
//     timeStart.setHours(8, 0, 0, 0)
//     const timeEnd = new Date(dates[i])
//     timeEnd.setDate(timeEnd.getDate() + 1)
//     timeEnd.setHours(0, 0, 0, 0)

//     //IDループ
//     const lastrow = sheetSalesmanId.getLastRow(); // 最後の列を取得
//     for (let idrow = 1; idrow  <= lastrow; idrow++) {
//       let eachId = sheetSalesmanId.getRange(idrow, 1).getValues();
//       Logger.log(eachId)
//       const events = CalendarApp.getCalendarById(eachId).getEvents(timeStart,timeEnd);// カレンダーから検索する予定を取得
      
//       //予定数ループ
//       for (let j = 0; j < events.length; j++) {
//         var title = events[j].getTitle();
//         var startTime = events[j].getStartTime();
//         var endTime = events[j].getEndTime();
//         var creators = events[j].getCreators();
//         var createdDate = events[j].getDateCreated();
//         var description = events[j].getDescription(); 


//         var applicantId = ""; // applicantIdの宣言と初期化を行う
//         // var match = description.match(/hatch\.apodori@gmail\.com\s+(\d+)/);
//         var match = description.match(/\d{8,}/);
//         if (match) {
//           applicantId = match[0];
//         }else{
//         }
        
//         var data = [eachId, title, startTime, endTime, creators, createdDate, applicantId, description];
//         //  description,
//         const eventData = []
//         eventData.push(data);
//         Logger.log(eventData)
//         Logger.log(data)

//         sheetLog.insertRowBefore(2);
//         sheetLog.getRange(2, 1, 1, data.length).setValues([data]);
//       }
//     }
//   }
// }




