function getClassroomData() {

  // 現在のスプレッドシート
  var spreadsheet = SpreadsheetApp.getActive();

  // **************************************************
  // 列クリア
  // **************************************************
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, commentsOnly: true, skipFilteredRows: true});

  // Classroom 一覧  
  var json = Classroom.Courses.list();

  // Classroom 数
  var cnt = json.courses.length;
  
  // トピック一覧用
  var jsonTopic = null;

  // ユーザ情報用
  var userProfile = null;

  // セル用
  var targetRange = null;

  // 行
  var row = 0;

  for( var i = 0; i < cnt; i++ ) {
    targetRange = spreadsheet.getRange('A' + (row + 1));
    targetRange.setValue(json.courses[i].id);
    targetRange = spreadsheet.getRange('B' + (row + 1));
    targetRange.setValue(json.courses[i].name);

    try {
      // ユーザ情報
      userProfile = Classroom.UserProfiles.get( json.courses[i].ownerId );
      targetRange = spreadsheet.getRange('E' + (row + 1));
      targetRange.setValue(userProfile.name.fullName);
      targetRange = spreadsheet.getRange('F' + (row + 1));
      targetRange.setValue(userProfile.emailAddress);
    }
    catch(e){
    }

    targetRange = spreadsheet.getRange('G' + (row + 1));
    targetRange.setValue(json.courses[i].creationTime);

    // トピック一覧
    jsonTopic = Classroom.Courses.Topics.list( json.courses[i].id );
    if ( jsonTopic.topic == null ) {
      row += 2;
    }
    else {
      try {
        for ( var j = 0; j < jsonTopic.topic.length; j++ ) {
          targetRange = spreadsheet.getRange('C' + (row +1));
          targetRange.setValue(jsonTopic.topic[j].topicId);
          targetRange = spreadsheet.getRange('D' + (row + 1));
          targetRange.setValue(jsonTopic.topic[j].name);
          row++;
        }
      }
      catch(e){
      }

      row++;
    }     

  }

}
