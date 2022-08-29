function getClassroomData() {

  // **************************************************
  // 現在のスプレッドシート
  // **************************************************
  var spreadsheet = SpreadsheetApp.getActive();

  // **************************************************
  // 列クリア
  // **************************************************
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, commentsOnly: true, skipFilteredRows: true});

  // **************************************************
  // Classroom 一覧  
  // **************************************************
  var json = Classroom.Courses.list();

  // **************************************************
  // Classroom 数
  // **************************************************
  var cnt = json.courses.length;
  
  // **************************************************
  // トピック一覧用
  // **************************************************
  var jsonTopic = null;

  // **************************************************
  // ユーザ情報用
  // **************************************************
  var userProfile = null;

  // **************************************************
  // セル用
  // **************************************************
  var targetRange = null;

  // **************************************************
  // 行
  // **************************************************
  var row = 0;

  for( var i = 0; i < cnt; i++ ) {
    // ************************************************
    // Classroom コースID
    // ************************************************
    targetRange = spreadsheet.getRange('A' + (row + 1));
    targetRange.setValue(json.courses[i].id);

    // ************************************************
    // コース名
    // ************************************************
    targetRange = spreadsheet.getRange('B' + (row + 1));
    targetRange.setValue(json.courses[i].name);

    try {
      // **********************************************
      // ユーザ情報
      // **********************************************
      userProfile = Classroom.UserProfiles.get( json.courses[i].ownerId );

      // **********************************************
      // ユーザ名
      // **********************************************
      targetRange = spreadsheet.getRange('E' + (row + 1));
      targetRange.setValue(userProfile.name.fullName);

      // **********************************************
      // メールアドレス
      // **********************************************
      targetRange = spreadsheet.getRange('F' + (row + 1));
      targetRange.setValue(userProfile.emailAddress);
    }
    catch(e){
    }

    // ************************************************
    // 実行時ログ
    // ************************************************
    Logger.log(json.courses[i].name + " : " + userProfile.name.fullName);

    targetRange = spreadsheet.getRange('G' + (row + 1));
    targetRange.setValue(json.courses[i].creationTime);

    // ************************************************
    // トピック一覧
    // ************************************************
    jsonTopic = Classroom.Courses.Topics.list( json.courses[i].id );
    if ( jsonTopic.topic == null ) {
      // **********************************************
      // トピックが全く無い場合は、２行進める
      // **********************************************
      row += 2;
    }
    else {
      // **********************************************
      // トピックの一覧を C列とD列に設定
      // **********************************************
      try {
        for ( var j = 0; j < jsonTopic.topic.length; j++ ) {
          // ******************************************
          // ID
          // ******************************************
          targetRange = spreadsheet.getRange('C' + (row + 1));
          targetRange.setValue(jsonTopic.topic[j].topicId);

          // ******************************************
          // トピック名称
          // ******************************************
          targetRange = spreadsheet.getRange('D' + (row + 1));
          targetRange.setValue(jsonTopic.topic[j].name);
          row++;

          // ******************************************
          // 実行時ログ
          // ******************************************
          Logger.log("    " + jsonTopic.topic[j].name);

        }
      }
      catch(e){
      }

      // **********************************************
      // 次のクラスと一行空ける
      // **********************************************
      row++;
    }     

  }

}
