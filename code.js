function onOpen(){
  var myMenu=[
    {name: "全取得(iPad)", functionName: "copyToOtherSht2"},
    {name: "全取得(AppleTV)", functionName: "copyToOtherSht3"},
    {name: "固定IPアドレス設定", functionName: "copyToOtherSht4"},
    {name: "全削除", functionName: "deleteALLSheets"},
    {name: "作業用シート生成", functionName: "generateWorkSheet"},
    {name: "Drive読込", functionName: "Tom_getFileListInFolder"},
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("コピー",myMenu); 
}

// 手入力、アプリ一覧、mobi
function copyToOtherSht0() {
  var sheetName = "★手入力";
  copyToOtherSht(sheetName,0);
  sheetName = "アプリ一覧";
  copyToOtherSht(sheetName,1);
  sheetName = "mobi";
  copyToOtherSht(sheetName,1);
}

// iPad作業工程①② 印刷用、iPad作業工程①②、iPad検査
function copyToOtherSht1() {
  var sheetName = "iPad作業工程①②";
  copyToOtherSht(sheetName,2);
  sheetName = "iPad作業工程①② 印刷用";
  copyToOtherSht(sheetName,3);
  sheetName = "iPad検査";
  copyToOtherSht(sheetName,5);
  sheetName = "iPad作業工程③(生徒)";
  copyToOtherSht(sheetName,6);
  sheetName = "iPad作業工程③(教員) ";
  copyToOtherSht(sheetName,7);
}

// iPad用作業シート全取得
function copyToOtherSht2() {
  var sheetName = "★手入力";
  copyToOtherSht(sheetName,0);
  sheetName = "アプリ一覧";
  copyToOtherSht(sheetName,1);
  sheetName = "mobi";
  copyToOtherSht(sheetName,1);
  sheetName = "iPad作業工程①②";
  copyToOtherSht(sheetName,2);
  sheetName = "iPad作業工程①② 印刷用";
  copyToOtherSht(sheetName,3);
  sheetName = "iPad検査";
  copyToOtherSht(sheetName,5);
  sheetName = "iPad作業工程③(生徒)";
  copyToOtherSht(sheetName,6);
  sheetName = "iPad作業工程③(教員) ";
  copyToOtherSht(sheetName,7);
}

// AppleTV用作業シート全取得(iPad作業シートが存在する前提)
function copyToOtherSht3() {
  var sheetName = "AppleTV作業工程";
  copyToOtherSht(sheetName,0);
  sheetName = "AppleTV作業工程 印刷用";
  copyToOtherSht(sheetName,0);
  sheetName = "AppleTV検査";
  copyToOtherSht(sheetName,0);
}

// 固定IPアドレス設定用作業シート全取得(iPad作業シートが存在する前提)
function copyToOtherSht4() {
  var sheetName = "iPad固定IP設定";
  copyToOtherSht(sheetName,0);
}

/**
 * 共通関数
 * sheetName:テンプレートから取得したいシート名
 * order:並び替えの順番
 *  */
function copyToOtherSht(sheetName,order) {
  var ssId = "スプレッドシートのIDを記入";

  var getSht = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);

  //コピーしたいシートをコピー先のスプレッドシートに移動
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var copySht = getSht.copyTo(activeSs);

  //コピーしたシート名を変更する
  copySht.setName(sheetName);
  if(order!=0) {
    var newSheet = activeSs.getSheetByName(sheetName).activate();
    activeSs.moveActiveSheet(order);
  }

}

/**
 * 削除除外シートor末尾のシート1つを残して、すべてのシートを削除する関数
 */
function deleteALLSheets() {
  //あらかじめ削除したくないシート名を記載してください。例「["シート１","シート5","シート10"]」
  const notDelSheet = ["★iPad初期設定","★iPad追加設定","★AppleTV初期設定","★AppleTV追加設定","★iPad固定IPアドレス設定"];
  // 現在アクティブなスプレッドシートを取得
  let mySheet = SpreadsheetApp.getActiveSpreadsheet();
  //取得したスプレッドシートにある全てのシートを配列として取得
  let sheetData = mySheet.getSheets();
  //末尾のシートを削除するかを決めるフラグ
  let flag =0;
  //forループでシートを削除する削除を行う
  for(i=0;i<sheetData.length;i++){
    //削除したくないシート存在しない場合、末尾のシートは削除せずスキップする
    if(flag ==0 && i == sheetData.length-1){
      break;
    }
    //削除対象から除外するシートにヒットした場合は、削除処理は行わず、フラグを立てる
    if(notDelSheet.indexOf(sheetData[i].getSheetName()) != -1){
      flag = 1;
    }else{  //削除除外シートではない場合は、削除処理を実行する
    mySheet.deleteSheet(sheetData[i]);
    }
  }
}

function generateWorkSheet() {

  // var id = Browser.inputBox("コピーしたいシートIDを入力");
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = ss.getId();
  // コピーするシートを呼び出し
  var sheet = DriveApp.getFileById(id);

  // 保存先のフォルダオブジェクトを取得
  var destfolderid = Browser.inputBox("コピー先のフォルダIDを入力");
  var destfolder = DriveApp.getFolderById(destfolderid);

  // シートに付与する文字を定義
  let alphabet = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

  // そのスプレッドシートのコピーを作成
  var filename;
  for (let i = 0; i < alphabet.length; i++) {
    filename = sheet.getName() + " " + alphabet[i];
    sheet.makeCopy(filename, destfolder);
  }
}

function Tom_getFileListInFolder() {
//フォルダのURLを入力してもらいIDを指定して、取得した情報を返す
  try{
    var folder_name = Browser.inputBox("Google DriveのフォルダURLを入れてください");
    var folder_id=folder_name.replace("https://drive.google.com/drive/folders/","");

    folder = DriveApp.getFolderById(folder_id);
    files = folder.getFiles();
    return files;

  }catch(e){
    Browser.msgBox(e);
  }

}

//同名シートを押し付けて更新するメソッド
function ForceOverWriteSheet(name) {
  try {
    //var sheet_name = "iPad作業工程①②";
    var sheet_name = name;
    var flies = Tom_getFileListInFolder();
    // 現在アクティブなスプレッドシートを取得
    var source = SpreadsheetApp.getActiveSpreadsheet();
    // そのスプレッドシートにあるコピー元シートを取得
    var sheet = source.getSheetByName(sheet_name);
    while(files.hasNext()) {
          var buff = files.next();
          var sheet_url = buff.getUrl();
          var sheet_id = sheet_url.replace("https://docs.google.com/spreadsheets/d/","").replace("/edit?usp=drivesdk","")
          // ID からスプレッドシートを取得
          var destination = SpreadsheetApp.openById(sheet_id);
          var delete_sheet = destination.getSheetByName(sheet_name);
          // 元々ある同名シートを削除
          destination.deleteSheet(delete_sheet);
          // 上記で取得したシートを、そのスプレッドシートにコピー
          sheet.copyTo(destination).setName(sheet_name); 
    };
   
    
  } catch(e) {
    Browser.msgBox(e);
  
  }
}
