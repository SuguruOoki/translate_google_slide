// ファイルコピー + ファイル名直し => 中身の翻訳(テキストとテーブルの内容 => translateに入れて返り値で置換)

/*
 *  @author Suguru oki <oki.suguru@hamee.co.jp>
 *  @param {string} copyFileName - コピー元のファイル名
 *  @param {string} original - 翻訳前の言語
 *  @param {string} translated - 翻訳後の言語
 *  @return {string}  copyFile.getId() - コピーしたファイルのID
 */
function fileCopy(copyFileName, original, translated) {
  
  const storeFolderId = "1RG7_p5y4gXXAD62xRTeL2OQbekFvXf0I"; // 格納先フォルダID
  const copyFolderId  = "1gAFvvIHq57zGqeknA5PFTP2LHRyAhrb9"; // コピー元フォルダID
  const storeFolder = DriveApp.getFolderById(storeFolderId);
  const copyFolder = DriveApp.getFolderById(copyFolderId);
  
  if (!copyFolder.getFilesByName(copyFileName).hasNext()) {
    Browser.msgBox(copyFileName + "がありません");
    return;
  }
 
  // コピー元ファイル
  var copyFile = copyFolder.getFilesByName(copyFileName).next();
  Logger.log(copyFile)
  var koreanFileName = LanguageApp.translate(copyFileName, original, translated);
  Logger.log(koreanFileName);
  var copiedFile = copyFile.makeCopy(koreanFileName + "(" + copyFileName + ")", storeFolder);
  if (!copiedFile) {
    return;
  }
  return copyFile.getId();
}

function updateShape(presentationID) {
  //  プレゼンテーションIDを指定し、プレゼンテーションを取得
  var presentation = Slides.Presentations.get(presentationID)
  //  1ページ目のスライドを取得
  var slide = presentation.slides[0]
  //  ページの要素を取得する
  var pageElements = slide["pageElements"]

  for(var i =0;i<pageElements.length;i++){
  //    ページ要素を取得
    var element = pageElements[i]
    var objectId = pageElements[i]["objectId"]
    if("shape" in element){
      // テキストボックスを取得
      var shape = element["shape"]
      // 中身のテキストを取得
      var content = shape["text"]["textElements"][1]["textRun"]["content"]

      if(content=="SampleText"){
        var text = "Replaced Text"
        var requests=[
          // テキスト削除リクエスト
          {
            "deleteText":{
              "objectId":objectId,
              "textRange":{
                "type":"ALL"
              }
            },
          },
          // テキスト挿入リクエスト
          {
            "insertText":{
              "objectId":objectId,
              "text":text,
              "insertionIndex":0
            }
          }

        ]
        Slides.Presentations.batchUpdate({"requests":requests},presentationID)
      }
    }
  }
}


function updateTable(presentationID) {
  //  プレゼンテーションIDを指定し、プレゼンテーションを取得
  var presentation = Slides.Presentations.get(presentationID)
  //  1ページ目のスライドを取得
  var slide = presentation.slides[0]
  //  ページの要素を取得する
  var pageElements = slide["pageElements"]

  for(var i =0;i<pageElements.length;i++){
    //    ページ要素を取得
    var element = pageElements[i]
    //    オブジェクトのIDを取得
    var objectId = pageElements[i]["objectId"]

    if("table"in element){
      var table =element["table"]
      // 置き換えるセルの行、列を指定
      var rowIndex = 1
      var columnIndex = 1
      // 指定したセルのテキストを取得
      var content = table["tableRows"][0]["tableCells"][rowIndex]["text"]["textElements"][columnIndex]["textRun"]["content"]
      var text = "AAAAA"
      var requests=[
        {
          "deleteText":{
            "objectId":objectId,
            "cellLocation":{
              "rowIndex":rowIndex,
              "columnIndex":columnIndex
            },
            "textRange":{
              "type":"ALL"
            }
          },
        },{
          "insertText":{
            "objectId":objectId,
            "cellLocation":{
              "rowIndex":rowIndex,
              "columnIndex":columnIndex
            },
            "text":text,
            "insertionIndex":0
          }
        }
      ]
      Slides.Presentations.batchUpdate({"requests":requests},presentationID)
    }
  }
}


var CLIENT_ID = '754354979392-qo8evfnfbprhqo49mb6ptnvsoc29bv2r.apps.googleusercontent.com';
var CLIENT_SECRET = 'ouWY5J0XUJnBPgqWpZDxVsWg';
var PRESENTATION_ID = '...';

// from https://mashe.hawksey.info/2015/10/setting-up-oauth2-access-with-google-apps-script-blogger-api-example/

function getService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService('slidesOauth')

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')


      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope('https://www.googleapis.com/auth/drive')

      // Below are Google-specific OAuth2 parameters.

      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getActiveUser().getEmail())

      // Requests offline access.
      .setParam('access_type', 'offline')

      // Forces the approval prompt every time. This is useful for testing,
      // but not desirable in a production application.
      .setParam('approval_prompt', 'force');
}

function authCallback(request) {
  var oauthService = getService();
  var isAuthorized = oauthService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function getSlideIds(presentationId) {
  var url = 'https://slides.googleapis.com/v1/presentations/' + presentationId;
  var options = {
    headers: {
      Authorization: 'Bearer ' + getService().getAccessToken()
    }
  };
}

function main () {

  const original = 'ja';
  const translated = 'ko';
  var copyFileName = '翻訳サンプル';
  
  var copiedFileId = fileCopy(copyFileName, original, translated);
  updateShape(copiedFileId);
  updateTable(copiedFileId);
}
