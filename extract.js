// https://qiita.com/kazinoue/items/1e8ed4aebfb5c3c886db

var HOURS_BEFORE = 12;

function createFirestore() {
  var email = "example@example.com";
  var key = "-----BEGIN PRIVATE KEY-----\nsomething-----END PRIVATE KEY-----\n";
  var projectId = "example";

  var firestore = FirestoreApp.getFirestore(email, key, projectId);
  return firestore;
}

function _createLabel(labelString) {
  labelDomain = GmailApp.getUserLabelByName(labelString);

  // 指定されたラベルが無い場合は作る
  if (labelDomain == null) {
    labelDomain = GmailApp.createLabel(labelString);
  }

  return labelDomain;
}

function myFunction() {
  var firestore = createFirestore();
  var label = GmailApp.getUserLabelByName("yamaha-booking");
  var doneLabel = _createLabel("firestored");
  var failLabel = _createLabel("fail");

  var after = parseInt(
    (new Date().getTime() - HOURS_BEFORE * 60 * 60 * 1000) / 1000
  );

  var searchTarget =
    "label:" +
    label.getName() +
    " after:" +
    after +
    " -label:" +
    doneLabel.getName();

  var myThreads = GmailApp.search(searchTarget, 0, 100);
  var myMsgs = GmailApp.getMessagesForThreads(myThreads);

  if (myThreads.length > 0) {
    for (var threadIndex = 0; threadIndex < myThreads.length; threadIndex++) {
      var mailBody = myMsgs[threadIndex][0].getPlainBody();

      myThreads[threadIndex].addLabel(doneLabel);

      var teacherMatch = mailBody.match(/【講師】\r\n([\s\S]*)\r\n【予/m);
      var teacher = teacherMatch ? teacherMatch[1].replace(/\s/, "") : "NA";
      console.log("teacher is:", teacher);

      var bookingIDMatch = mailBody.match(/【予約番号】\r\n([\s\S]*)\r\n【コ/m);
      var bookingID = bookingIDMatch ? bookingIDMatch[1] : "NA";
      console.log("bookingID is:", bookingID);

      var bookingDateMatch = mailBody.match(
        /【予約日】\r\n([\s\S]*)\r\n【\s時/
      );
      var bookingDate = bookingDateMatch[1] ? bookingDateMatch[1] : "NA";
      console.log("bookingDate is:", bookingDate);

      var bookingTimeMatch = mailBody.match(
        /【 時  間 】\r\n([\s\S]*)\r\n【サ/m
      );
      var bookingTime = bookingTimeMatch[1] ? bookingTimeMatch[1] : "NA";
      console.log("bookingTime is:", bookingTime);

      var skypeIDMatch = mailBody.match(
        /【サービス時に利用する\[Skype名\]をご記入ください。】\r\n([\s\S]*)\r\n【こ/m
      );
      var skypeID = skypeIDMatch ? skypeIDMatch[1] : "NA";
      console.log("skypeID is:", skypeID);

      var songsMatch = mailBody.match(
        /【このレッスンでみてほしい曲名や、家庭学習でお困りのことについてご記入ください。】\r\n([\s\S]*)\r\n【そ/m
      );
      var songs = songsMatch ? songsMatch[1].replace(/\r\n/g, "") : "NA";
      console.log("songs are:", songs);

      var messagesMatch = mailBody.match(
        /【その他、担当講師へのメッセージなどございましたらご記入ください。】\r\n([\s\S]*?)\r\n-/m
      );
      messages = messagesMatch ? messagesMatch[1].replace(/\r\n/g, "") : "NA";
      console.log("messages are:", messages);

      var kidNameMatch = mailBody.match(
        /【お子さまのお名前（下のお名前）】\r\n([\s\S]*)\r\n【お/m
      );
      var kidName = kidNameMatch ? kidNameMatch[1] : "NA";
      console.log("kidName is:", kidName);

      var courseMatch = mailBody.match(
        /【お通いのコース名】\r\n([\s\S]*?)\r\n-/m
      );
      var course = courseMatch ? courseMatch[1].replace(/\r\n/g, "") : "NA";
      console.log("course is:", course);

      var dateMatch = bookingDate.match(/(\d+)\/(\d+)\/(\d+)/);
      var year = dateMatch[1];
      var month = dateMatch[2] - 1;
      var date = dateMatch[3];

      var timeMatch = bookingTime.match(/(\d+)\:(\d+)/);
      var hour = timeMatch[1];
      var min = timeMatch[2];

      var lessonDateObject = new Date(year, month, date, hour, min, 0);
      console.log(lessonDateObject.toLocaleString());

      var data = {
        lessonDate: lessonDateObject.getTime(),
        bookingID: bookingID,
        bookingDate: bookingDate,
        bookingTime: bookingTime,
        skypeID: skypeID,
        songs: songs,
        messages: messages,
        kidName: kidName,
        course: course,
      };

      try {
        firestore.createDocument(
          teacher + "_schedules/" + data.bookingID,
          data
        );
      } catch (err) {
        let dataString = "";
        const dataKeys = Object.keys(data);
        dataKeys.map(function (prop) {
          dataString += prop + ": " + data[prop] + " ";
        });

        myThreads[threadIndex].addLabel(failLabel);
        MailApp.sendEmail(
          "example@gmail.com",
          "エラー発生",
          err + "\n\n\n\n" + dataString
        );
      }
    }
  }
}
