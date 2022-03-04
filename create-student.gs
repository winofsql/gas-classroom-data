function createStudent() {

	var CourceId = "コースID";

	// 選択したシートを対象とします
	var spreadsheet = SpreadsheetApp.getActive();

	// 行番号
	var i = 1;

	while (true) {

		// A 列の先頭から順に取得して、空白になったら終了
		var targetRange = spreadsheet.getRange('A' + i);
		var cellWork = targetRange.getValue().toString();
		if (cellWork != '') {

			// 生徒作成用の JSON
			var json = {
				"userId": cellWork + "@ドメイン"
			};

			// 生徒を追加
			// ( 招待済でも確定します )
			try {
				Classroom.Courses.Students.create(json, CourceId);

			}
			catch(e) {
				GmailApp.sendEmail("メールアドレス", "Classroom 生徒登録エラー", JSON.stringify(json) + "\r\n" + e.message );

			}

			i++;

		}
		else {
			break;
		}
	}
}
