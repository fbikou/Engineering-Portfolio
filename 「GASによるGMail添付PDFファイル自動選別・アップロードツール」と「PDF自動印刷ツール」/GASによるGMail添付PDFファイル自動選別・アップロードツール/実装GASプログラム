//main関数(検索ワード格納+新着メールのPDFファイル添付有無判定+他機能)
function PDFSelectDownload() {

  //検索ワード格納用
  //＊ここ[openByURL()の()内]にスプレッドシートのURLを貼り付けてください
  const ss=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1XsA4uXPuHkwhIWmR1nKCrhMmKyoFl0aXl3T2nUochYQ/edit#gid=0');
  const WordsSheet=ss.getSheetByName('検索ワード一覧');
  let HitWords=[];
  //新着・ファイル添付有り・メアド条件設定用
  let condition='';
  const addresses=['メアド'];
  let SearchThreads = null;
  let SearchMails=null;

  //検索ワード格納
  for(let i=2;i<102;i++){
    if(WordsSheet.getRange(i,1).isBlank()===false){
      HitWords.push(WordsSheet.getRange(i,1).getValue());
    }
  }
  for(let j=0;j<addresses.length;j++){
    //対象メールの再設定
    condition='has:attachment from:'+addresses[j];//添付あり・メアド条件設定
    SearchThreads=GmailApp.search(condition,0,500);
    SearchMails= GmailApp.getMessagesForThreads(SearchThreads);
    //メール毎の確認
    for (let k in SearchMails) {
      for (let l in SearchMails[k]) {
        ForOneMail(SearchMails[k][l],HitWords);
      }
    }
  }
}
//メール単体機能
function ForOneMail(mail,HitWords){
  //添付ファイル・拡張子格納用
  let attachment=null;
  let AttachmentName='';
  let extension='';
  //GoogleDriveのダウンロードファイル名
  const FolderName='該当PDFファイル一覧';
  //該当PDF判定結果格納用
  let HitPDFJudge=[];

  //メールに✰が付いてなかったら
  if(!mail.isStarred()){
    //添付ファイル格納
    attachment=mail.getAttachments();
    //attachmentは配列の為, 以降, for文を用いる
    for (let m in attachment){
      AttachmentName=attachment[m].getName()
      extension=AttachmentName.substring(AttachmentName.indexOf('.')+1,AttachmentName.length+1);
      if(extension==='pdf'){
        HitPDFJudge=PDFJudge(FolderName,attachment[m],HitWords)
        if(HitPDFJudge[0]===false){
          Drive.Files.remove(HitPDFJudge[1]);
        }
      }
    }
    //メールに✰を付ける
    mail.star();
  }
}
//PDF読み込み・判定関数
function PDFJudge(FolderName,attachment,HitWords) {

  let JudgeResult=false;
  const option = {
    'ocr': true,        // OCRを行うかの設定
    'ocrLanguage': 'ja',// OCRを行う言語の設定
  }
  const resource = {
    title: attachment.getName()
  };
  const PDFid=DriveUp(FolderName,attachment).getId();
  const image = Drive.Files.copy(resource, PDFid, option);   // 指定したファイルをコピー
  const text = DocumentApp.openById(image.id).getBody().getText();  // コピー先ファイルのOCRのデータを取得
  Drive.Files.remove(image.id);
  let HitCount=0;

  for(let n in HitWords){
    if(text.indexOf(HitWords[n])!=-1){
      HitCount++;
    }
  }
  if(HitCount>=3){
    JudgeResult=true;
  }
  return [JudgeResult,PDFid];
}
//GoogleDriveに自動アップロードする関数
function DriveUp(FolderName,attachment) {

  let folder=null;
  
  const folderIterator = DriveApp.getRootFolder().getFoldersByName(FolderName);

  if (folderIterator.hasNext()) {
    // 存在する場合
    folder = folderIterator.next();
  } else {
    // 存在しない場合
    folder = DriveApp.getRootFolder().createFolder(FolderName);
  }
  //添付ファイルを指定フォルダに格納
  return folder.createFile(attachment);
}
