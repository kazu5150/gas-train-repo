//◆◆◆業務管理Spreadsheetを取得◆◆◆
const ss = SpreadsheetApp.getActiveSpreadsheet();


//◆◆◆案件を取得◆◆◆
class JobClass{
  constructor(jobRecord){
    [this.jobFolder,
    this.jobId,
    this.jobNum,
    this.jobCliant,
    this.jobCliantPerson,
    this.jobName,
    this.jobDetail,
    this.jobGyoumuPerson,
    this.jobEigyouPerson,
    this.jobFirstShippingDate,
    this.jobRemark,
    this.jobLeftovers,
    this.jobReport,
    this.jobJobFolderId,
    this.jobJobNote1,
    this.jobJobNote2,
    this.jobJobNote3,
    this.jobYearMonth
    ] = jobRecord;
  }
}



const jobSheet = ss.getSheetByName("案件");  //案件シートを取得
const rngJobSheet = jobSheet.getDataRange();  //案件シートのRangeを取得
const aryJobSheet = rngJobSheet.getValues();  //案件を配列で取得
const numJobSheetLastRow = rngJobSheet.getLastRow();  //案件シート最終行
const numJobSheetLastCol = rngJobSheet.getLastColumn();  //案件シート最終列
const rngJobSheetCurrentCell = jobSheet.getCurrentCell();  //案件シートのアクティブセルを取得
const numJobSheetCurrentRow = rngJobSheetCurrentCell.getRow();  //案件シートのアクティブセル行番号
const numJobSheetCurrentCol = rngJobSheetCurrentCell.getColumn();  //案件シートのアクティブセル列番号


// const testObject = new JobClass(aryJobSheet[numJobSheetCurrentRow-1]);
// Logger.log(testObject.jobYearMonth);



//◆◆◆予定を取得◆◆◆
class ScheduleClass{
  constructor(schduleRecord){
    [this.scheduleStatus,
    this.scheduleSchduleId,
    this.scheduleJobId,
    this.scheduleJobNum,
    this.scheduleCliant,
    this.scheduleCliantPerson,
    this.scheduleJobName,
    this.scheduleJobDetail,
    this.scheduleQuantity,
    this.scheduleDmSticker,
    this.schedulePlace,
    this.scheduleArrivalDate,
    this.scheduleShippingDate,
    this.scheduleShippingTime,
    this.scheduleShippngMethod,
    this.scheduleGyoumuPerson,
    this.scheduleEigyouPerson,
    this.scheduleRemark,
    this.scheduleCheckSheet,
    this.schedulePlTag,
    this.scheduleJobFolderId
    ] = schduleRecord;
  }
}

const scheduleSheet = ss.getSheetByName("予定");  //予定シートを取得
const rngscheduleSheet = scheduleSheet.getDataRange();  //予定シートのRangeを取得
const aryScheduleSheet = rngscheduleSheet.getValues();  //予定を配列で取得
const numScheduleSheetLastRow =　rngscheduleSheet.getLastRow();  //予定シート最終行
const numScheduleSheetLastCol = 　rngscheduleSheet.getLastColumn();  //予定シート最終列
const rngScheduleSheetCurrentCell = scheduleSheet.getCurrentCell();  //予定シートのアクティブなセルを取得
const numScheduleSheetCurrentRow = rngScheduleSheetCurrentCell.getRow();  //予定シートのアクティブセル行番号
const numScheduleSheetCurrentCol = rngScheduleSheetCurrentCell.getColumn();  //予定シートのアクティブセル列番号




//★★担当者フォルダーを取得(共有ドライブ内)★★
const matsuzawaFolder = DriveApp.getFolderById("1TNjwuundYH-rPoGpAIwJ81Yv0unhGzzK") ;
const yabutaFolder =  DriveApp.getFolderById("1CmeqErA5brmCgpZ7p6_TJHp21LeqLrjx") ;
const yakoFolder =  DriveApp.getFolderById("1XOPcHIO3Sx1Y0hkESk8a8t-CGLd2vNs_") ;
const uchiyamaFolder =  DriveApp.getFolderById("1APDFdkCr8hWdzfzoIIsTDC8CrO7DMe77") ;
const kitamotoFolder =  DriveApp.getFolderById("1C0aPkfKJ67p1R7J96Adouhor2Q08wywE") ;
const hashimotoFolder =  DriveApp.getFolderById("1mxfxNqCMh2p6wRRqNOGqkedwfTiHe5Qh") ;
const ohkiFolder =  DriveApp.getFolderById("1BNuHDZFeAaDJDajDhljcUjhyrzYIkk9Z") ;
const iwamotoFolder =  DriveApp.getFolderById("1TMu4Dss9xCIPEu_norWd57AUdUKunlay") ;
const kakizoeFolder =  DriveApp.getFolderById("1DzV94l-205WsoAblCAA6M9LvxgmAn9Aw") ;
const yamadaFolder =  DriveApp.getFolderById("1DImewQKUVJkilVgLgUMvjBZq0EnHZ3dY") ;








//★★作業報告書テンプレートを取得★★
const templateReport = SpreadsheetApp.openById("1XVPPAvV2y-yuQ6px2dQRHfMLdY1GegVqWI-g4ifS_pw");
const templateReportFile = DriveApp.getFileById("1XVPPAvV2y-yuQ6px2dQRHfMLdY1GegVqWI-g4ifS_pw");

//Logger.log(templateReport);

const rngReportCliantName = templateReport.getRange("E3");
const rngReportCliantPerson = templateReport.getRange("E4");
const rngReportJobName = templateReport.getRange("E5");
const rngReportJobNumber = templateReport.getRange("E6");
const rngReportJobId = templateReport.getRange("M3");
const rngReportGyoumuPerson = templateReport.getRange("M4");
const rngReporEigyouPerson = templateReport.getRange("M5");


//★★PL札テンプレートを取得★★
const templatePltag = DriveApp.getFileById("1SLkU7DOYP4n3LRwatDI-HQ-QFnPwMzHDXWteeYyhmdg");  //局出、納品、引取、他　用のPL札テンプレート
const templatePltagDm = DriveApp.getFileById("1kOLj99E3Wr0mP4CnKMrX1k7UXkqAHSCkkd-nQlSslIs");  //クロネコDM便　用のPL札テンプレート

// const ssTemplatePltag = SpreadsheetApp.openById("1SLkU7DOYP4n3LRwatDI-HQ-QFnPwMzHDXWteeYyhmdg");
// const templatePltagDm = templatePltag.getSheetByName("DM便");
// const templatePltagOtherDm = templatePltag.getSheetByName("D局出･納品･引取･地区宅便");






//★★★★★　Spreadsheetを開いたら　案件シートの最終行の指定のセルをアクティブにする　★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function setActhiveCel(){
   const activeCell = jobSheet.getRange(numJobSheetLastRow+1,3,1,1);
  jobSheet.setActiveRange(activeCell);

}





//★★★★★　担当者フォルダー内に「案件フォルダー」を作成　★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function createJobFolder() {
  //担当者フォルダー内に新規案件フォルダーを作成

  const jobFolderUrl = jobSheet.getRange(numJobSheetCurrentRow,1,1,1).getValue();

  if (jobFolderUrl){
    Logger.log("既に案件フォルダーが存在します");
    //Browser.msgBox("既に案件フォルダーが存在します");
  }else{
    const objJob = new JobClass(aryJobSheet[numJobSheetCurrentRow-1]);

    const date = aryJobSheet[numJobSheetCurrentRow-1][numFirstShippingDateCol]
    const formatDate = Utilities.formatDate(date,"Asia/Tokyo","yyyy/MM/dd")
    const jobFolderName = formatDate +" "+ objJob.jobCliant +"様 "+"【"+objJob.jobName+"】"    //+" Jid："+objJob.jobId  フォルダー名にjobIdを付与したいが・・

    switch(objJob.jobGyoumuPerson){
      case "松澤":
        const jobFolderMatsuzawa = matsuzawaFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderMatsuzawa.getUrl());
        Logger.log(jobFolderMatsuzawa.getId())
        const jobFolderMatsuzawaId = jobFolderMatsuzawa.getId()
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderMatsuzawaId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:matsuzawa@sunprompt.com","松澤")')  //業務担当者欄にGmailリンクを追加

        // 参考コード　 sheet.getRange(1,1).setFormula('=HYPERLINK("' + url + '","テスト")'); 

      break;

      case "薮田":
        const jobFolderYabuta = yabutaFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderYabuta.getUrl());
        Logger.log(jobFolderYabuta.getId())
        const jobFolderYabutaId = jobFolderYabuta.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderYabutaId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:yabuta@sunprompt.com","薮田")')   //業務担当者欄にGmailリンク
      break;

      case "八子":
        const jobFolderYako = yakoFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderYako.getUrl());
        Logger.log(jobFolderYako.getId())
        const jobFolderYakoId = jobFolderYako.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderYakoId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:yako@sunprompt.com","八子")')   //業務担当者欄にGmailリンク
      break;

      case "めぐみ":
        const jobFolderUchiyama = uchiyamaFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderUchiyama.getUrl());
        Logger.log(jobFolderUchiyama.getId())
        const jobFolderUchiyamaId = jobFolderUchiyama.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderUchiyamaId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:uchiyama@sunprompt.com","めぐみ")') //業務担当者欄にGmailリンク
      break;

      case "北本":
        const jobFolderKitamoto = kitamotoFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderKitamoto.getUrl());
        Logger.log(jobFolderKitamoto.getId())
        const jobFolderKitamotoId = jobFolderKitamoto.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderKitamotoId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:kitamoto@sunprompt.com","北本")') //業務担当者欄にGmailリンク
      break;

      case "橋本":
        const jobFolderHashimoto = hashimotoFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderHashimoto.getUrl());
        Logger.log(jobFolderHashimoto.getId())
        const jobFolderHashimotoId = jobFolderHashimoto.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderHashimotoId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:hashimoto@sunprompt.com","橋本")')  //業務担当者欄にGmailリンク
      break;

      case "黄木":
        const jobFolderOhki = ohkiFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderOhki.getUrl());
        Logger.log(jobFolderOhki.getId())
        const jobFolderOhkiId = jobFolderOhki.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderOhkiId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:ohki@sunprompt.com","黄木")')  //業務担当者欄にGmailリンク
      break;

      case "岩本":
        const jobFolderIwamoto = iwamotoFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderIwamoto.getUrl());
        Logger.log(jobFolderIwamoto.getId())
        const jobFolderIwamotoId = jobFolderIwamoto.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderIwamotoId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:iwamoto@sunprompt.com","岩本")') //業務担当者欄にGmailリンク
      break;

      case "柿添":
        const jobFolderKakizoe = kakizoeFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderKakizoe.getUrl());
        Logger.log(jobFolderKakizoe.getId())
        const jobFolderKakizoeId = jobFolderKakizoe.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderKakizoeId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:kakizoe@sunprompt.com","柿添")')  //業務担当者欄にGmailリンク
      break;

      case "山田":
        const jobFolderYamada = yamadaFolder.createFolder(jobFolderName);
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderCol+1,1,1).setValue(jobFolderYamada.getUrl());
        Logger.log(jobFolderYamada.getId())
        const jobFolderYamadaId = jobFolderYamada.getId();
        jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).setValue(jobFolderYamadaId);
        //jobSheet.getRange(numJobSheetCurrentRow,numGyoumuPersonCol+1,1,1).setFormula('=hyperlink("mailto:yamada@sunprompt.com","山田")')  //業務担当者欄にGmailリンク        
      break;
    }
  } 
}





//★★★★★　案件ID　予定IDをキーにDrive内のファイルを検索　★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
//検索クエリに合致するすべてのファイルを取得するプログラム
function DriveAppSearchFilesById() {

  //const jobId = aryJobSheet[numJobSheetCurrentRow][numJobIdCol];
  const jobId = aryJobSheet[2][0];

  Logger.log(jobId);

  //var str = "見積書";
  var params = "title contains " + "'" + jobId + "'"


  const searchedFiles = DriveApp.searchFiles(params);
  while (searchedFiles.hasNext()) {
    const file = searchedFiles.next();
    const fileName = file.getName();
    Logger.log(fileName);
    const fileUrl = file.getUrl();
    Logger.log(fileUrl);
  }
}























//★★★★★★メニュー作成★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function onOpen(){

  const activeSheet = ss.getActiveSheet().getName();
  Logger.log(activeSheet);


  const ui = SpreadsheetApp.getUi();

  const menu1 = ui.createMenu("【◆案件◆】ﾒﾆｭｰ");
  menu1.addItem("①案件を予定表に追加","copyValueFromJobSheetToSchduleSheet");
  menu1.addItem("②作業報告書を作成","createReportVer4");
  menu1.addItem("【作成中】案件ファイルを探す","DriveAppSearchFilesById");

  const menu2 = ui.createMenu("【◆予定◆】ﾒﾆｭｰ");
  menu2.addItem("①日付順にソートする","createBordar");
  menu2.addItem("②PL札を作成","createPltag");


  menu1.addToUi();
  menu2.addToUi();

}








//★★★★★★★予定表からPL札を作成★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function createPltag(){

  const objSchedule = new ScheduleClass(aryScheduleSheet[numScheduleSheetCurrentRow-1]);
  const jobFolderId = aryScheduleSheet[numScheduleSheetCurrentRow-1][numScheduleJobFolderIdCol]; 
  Logger.log(jobFolderId);


  //★★必要な値を取得★★
  // const rngScheduleSheet = scheduleSheet.getRange(1,1,numScheduleSheetLastRow,numScheduleSheetLastCol).getValues();

  // const gyoumu = rngScheduleSheet[numScheduleSheetCurrentRow-1][numScheduleGyoumuPersonCol];  //業務担当者名を取得
  // const date = rngScheduleSheet[numScheduleSheetCurrentRow-1][numScheduleShippingDateCol];  //出荷日を取得
  // const jobName = rngScheduleSheet[numScheduleSheetCurrentRow-1][numscheduleJobNameCol];  //案件名を取得
  // const scheduleId = rngScheduleSheet[numScheduleSheetCurrentRow-1][numScheduleIdCol];  //予定IDを取得
  // const method = rngScheduleSheet[numScheduleSheetCurrentRow-1][numShippngMethodCol];  //出荷方法を取得
  // const jobId = rngScheduleSheet[numScheduleSheetCurrentRow-1][numScheduleJobIdCol];  //案件IDを取得


  //★★PL札作成に必要な値を取得★★
  function setValuesTemplatePltagFile(){
    ssPltag.getRange("S2").setValue(objSchedule.scheduleSchduleId);
    ssPltag.getRange("G18").setValue(objSchedule.scheduleGyoumuPerson);
    ssPltag.getRange("G8").setValue(objSchedule.scheduleJobName);
    ssPltag.getRange("G13").setValue(objSchedule.scheduleShippingDate);
    ssPltag.getRange("C3").setValue(objSchedule.scheduleShippngMethod);

    //↓↓↓↓出荷日を構造化データで取得
    var formatDate = Utilities.formatDate(ssPltag.getRange("G13").getValue(),"Asia/Tokyo","yyyy/MM/dd")

    //↓↓↓↓PL作成したPL札に名前を付ける
    var plTagName = 
  　  "【PL札】　"+formatDate +" "+ 
      ssPltag.getRange("C3").getValue() + 
      "【"+ssPltag.getRange("G8").getValue()+"】"+
      "Sid-"+ssPltag.getRange("S2").getValue()+" "+
      "Jid-"+objSchedule.scheduleJobId;
 
    ssPltag.rename(plTagName)
 
  }


  switch (objSchedule.scheduleGyoumuPerson){
    case "松澤":
    if (objSchedule.scheduleShippngMethod.match(/DM/)) {
      const matsuzawaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyMatsuzawa = templatePltagDm.makeCopy("松澤　PL札", matsuzawaJobFolder);
      const templatePltagDmCopyMatsuzawaId = templatePltagDmCopyMatsuzawa.getId(); 
      const templatePltagDmCopyMatsuzawaUrl = templatePltagDmCopyMatsuzawa.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyMatsuzawaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyMatsuzawaUrl); 
    }else{
      const matsuzawaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyMatsuzawa = templatePltag.makeCopy("松澤　PL札", matsuzawaJobFolder);
      const templatePltagCopyMatsuzawaId = templatePltagCopyMatsuzawa.getId(); 
      const templatePltagCopyMatsuzawaUrl = templatePltagCopyMatsuzawa.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyMatsuzawaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyMatsuzawaUrl); 
    }
      break;
    
    case "薮田":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const yabutaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyYabuta = templatePltagDm.makeCopy("薮田　作業報告書", yabutaJobFolder);
      const templatePltagDmCopyYabutaId = templatePltagDmCopyYabuta.getId(); 
      const templatePltagDmCopyYabutaUrl = templatePltagDmCopyYabuta.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyYabutaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyYabutaUrl); 
    }else{
      const yabutaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyYabuta = templatePltag.makeCopy("薮田　作業報告書", yabutaJobFolder);
      const templatePltagCopyYabutaId = templatePltagCopyYabuta.getId(); 
      const templatePltagCopyYabutaUrl = templatePltagCopyYabuta.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyYabutaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyYabutaUrl); 
    }
    break;

    case "八子":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
        const yakoJobFolder = DriveApp.getFolderById(jobFolderId);
        const templatePltagDmCopyYako = templatePltagDm.makeCopy("八子　作業報告書", yakoJobFolder);
        const templatePltagDmCopyYakoId = templatePltagDmCopyYako.getId(); 
        const templatePltagDmCopyYakoUrl = templatePltagDmCopyYako.getUrl();
        var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyYakoId);
        setValuesTemplatePltagFile();
        scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyYakoUrl); 
      }else{
        const yakoJobFolder = DriveApp.getFolderById(jobFolderId);
        const templatePltagCopyYako = templatePltag.makeCopy("八子　作業報告書", yakoJobFolder);
        const templatePltagCopyYakoId = templatePltagCopyYako.getId(); 
        const templatePltagCopyYakoUrl = templatePltagCopyYako.getUrl();
        var ssPltag = SpreadsheetApp.openById(templatePltagCopyYakoId);
        setValuesTemplatePltagFile();
        scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyYakoUrl); 
      }
    break;

    case "めぐみ":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const uchiyamaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyUchiyama = templatePltagDm.makeCopy("内山　作業報告書", uchiyamaJobFolder);
      const templatePltagDmCopyUchiyamaId = templatePltagDmCopyUchiyama.getId(); 
      const templatePltagDmCopyUchiyamaUrl = templatePltagDmCopyUchiyama.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyUchiyamaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyUchiyamaUrl); 
    }else{
      const uchiyamaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyUchiyama = templatePltag.makeCopy("内山　作業報告書", uchiyamaJobFolder);
      const templatePltagCopyUchiyamaId = templatePltagCopyUchiyama.getId(); 
      const templatePltagCopyUchiyamaUrl = templatePltagCopyUchiyama.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyUchiyamaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyUchiyamaUrl);
    } 
    break;


    case "北本":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const kitamotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyKitamoto = templatePltagDm.makeCopy("北本　作業報告書", kitamotoJobFolder);
      const templatePltagDmCopyKitamotoId = templatePltagDmCopyKitamoto.getId(); 
      const templatePltagDmCopyKitamotoUrl = templatePltagDmCopyKitamoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyKitamotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyKitamotoUrl); 
    }else{
      const kitamotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyKitamoto = templatePltag.makeCopy("北本　作業報告書", kitamotoJobFolder);
      const templatePltagCopyKitamotoId = templatePltagCopyKitamoto.getId(); 
      const templatePltagCopyKitamotoUrl = templatePltagCopyKitamoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyKitamotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyKitamotoUrl); 
    }
    break;

    case "黄木":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const ohkiJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyOhki = templatePltagDm.makeCopy("北本　作業報告書", ohkiJobFolder);
      const templatePltagDmCopyOhkiId = templatePltagDmCopyOhki.getId(); 
      const templatePltagDmCopyOhkiUrl = templatePltagDmCopyOhki.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyOhkiId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyOhkiUrl); 
    }else{
      const ohkiJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyOhki = templatePltag.makeCopy("北本　作業報告書", ohkiJobFolder);
      const templatePltagCopyOhkiId = templatePltagCopyOhki.getId(); 
      const templatePltagCopyOhkiUrl = templatePltagCopyOhki.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyOhkiId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyOhkiUrl); 
    }
    break;

    case "橋本":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const hashimotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyHashimoto = templatePltagDm.makeCopy("橋本　作業報告書", hashimotoJobFolder);
      const templatePltagDmCopyHashimotoId = templatePltagDmCopyHashimoto.getId(); 
      const templatePltagDmCopyHashimotoUrl = templatePltagDmCopyHashimoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyHashimotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyHashimotoUrl); 
    }else{
      const hashimotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyHashimoto = templatePltag.makeCopy("橋本　作業報告書", hashimotoJobFolder);
      const templatePltagCopyHashimotoId = templatePltagCopyHashimoto.getId(); 
      const templatePltagCopyHashimotoUrl = templatePltagCopyHashimoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyHashimotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyHashimotoUrl); 
    }
    break;

    case "岩本":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const iwamotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyIwamoto = templatePltagDm.makeCopy("岩本　作業報告書", iwamotoJobFolder);
      const templatePltagDmCopyIwamotoId = templatePltagDmCopyIwamoto.getId(); 
      const templatePltagDmCopyIwamotoUrl = templatePltagDmCopyIwamoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyIwamotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyIwamotoUrl); 
    }else{
      const iwamotoJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyIwamoto = templatePltag.makeCopy("岩本　作業報告書", iwamotoJobFolder);
      const templatePltagCopyIwamotoId = templatePltagCopyIwamoto.getId(); 
      const templatePltagCopyIwamotoUrl = templatePltagCopyIwamoto.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyIwamotoId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyIwamotoUrl); 
    }
    break;

    case "柿添":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const kakizoeJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyKakizoe = templatePltagDm.makeCopy("柿添　作業報告書", kakizoeJobFolder);
      const templatePltagDmCopyKakizoeId = templatePltagDmCopyKakizoe.getId(); 
      const templatePltagDmCopyKakizoeUrl = templatePltagDmCopyKakizoe.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyKakizoeId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyKakizoeUrl); 
    }else{
      const kakizoeJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyKakizoe = templatePltag.makeCopy("柿添　作業報告書", kakizoeJobFolder);
      const templatePltagCopyKakizoeId = templatePltagCopyKakizoe.getId(); 
      const templatePltagCopyKakizoeUrl = templatePltagCopyKakizoe.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyKakizoeId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyKakizoeUrl); 
    }
    break;

    case "山田":
    if ( objSchedule.scheduleShippngMethod.match(/DM/)) {
      const yamadaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagDmCopyYamada = templatePltagDm.makeCopy("山田　作業報告書", yamadaJobFolder);
      const templatePltagDmCopyYamadaId = templatePltagDmCopyYamada.getId(); 
      const templatePltagDmCopyYamadaUrl = templatePltagDmCopyYamada.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagDmCopyYamadaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagDmCopyYamadaUrl); 
    }else{
      const yamadaJobFolder = DriveApp.getFolderById(jobFolderId);
      const templatePltagCopyYamada = templatePltag.makeCopy("山田　作業報告書", yamadaJobFolder);
      const templatePltagCopyYamadaId = templatePltagCopyYamada.getId(); 
      const templatePltagCopyYamadaUrl = templatePltagCopyYamada.getUrl();
      var ssPltag = SpreadsheetApp.openById(templatePltagCopyYamadaId);
      setValuesTemplatePltagFile();
      scheduleSheet.getRange(numScheduleSheetCurrentRow,numSchedulePlTagCol+1,1,1).setValue(templatePltagCopyYamadaUrl); 
    }
    break;
  }
}











//◆◆◆PDFを格納するフォルダーをGWS共有フォルダーにするとうまく動かない。　別途チームフォルダーを扱う設定が必要の様だ◆◆◆
/**url "https://officeforest.org/wp/2019/12/11/google-apps-script%E3%81%A7%E5%85%B1%E6%9C%89%E3%83%89%E3%83%A9%E3%82%A4%E3%83%96%E3%82%92%E3%82%B3%E3%83%B3%E3%83%88%E3%83%AD%E3%83%BC%E3%83%AB%E3%81%99%E3%82%8B/"
*/
//★★★★★★PDF（外注ノート）をフォルダーにUPすると、OCRを読み取りファイル名を自動で付与★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
//参照URL　https://www.iehohs.com/gas-ocr/
function pdf_to_text() {

  //★★外注ノートフォルダーIDを取得★★
  const ocrFolderID = '1k_zdEGV2GuwoL64jCJGq9_lXP-5nCgLt'; 
  let files = DriveApp.getFolderById(ocrFolderID).getFiles();   // OCR対象フォルダに入っているファイルを取得
  //Logger.log(files);
 
  let option = {
    'ocr': true,        // OCRを行うかの設定
    'ocrLanguage': 'ja',// OCRを行う言語の設定
  }
  while(files.hasNext()){     // 取得したファイルを１件ずつ処理
    let file = files.next();  // ファイル単体を取得
    subject = file.getName(); // ファイル名を取得
    Logger.log(subject);
    let resource = {
      title: subject
    };
    let image = Drive.Files.copy(resource, file.getId(), option);   // 指定したファイルをコピー
    Logger.log(image.id);
    Logger.log(typeof(image));
    let text = DocumentApp.openById(image.id).getBody().getText();  // コピー先ファイルのOCRのデータを取得
    //file.setName("【外注ノート】"+text.slice(0,50))
    // OCR後のデータ削除を行う場合はコメントアウトを外す
    //Drive.Files.remove(image.id);
    Logger.log(text.split("/")[0]); //外注ノート
    Logger.log(text.split("/")[1]); //案件ID
    Logger.log(text.split("/")[2]); //予定ID
    Logger.log(text.split("/")[3]); //顧客名
    Logger.log(text.split("/")[4]); //案件名
    Logger.log(text.split("/")[5]); //日付
    Logger.log(text.split("/")[6]); //
    Logger.log(text.split("/")[7]); //
    Logger.log(text.split("/")[8]); //
    
    const gatJobId = text.split("/")[1]; //案件ID
    const gatscheduleId = text.split("/")[2]; //予定ID
    const gatCliantName = text.split("/")[3]; //顧客名
    const gatJobName = text.split("/")[4]; //案件名
    const gatDate = text.split("/")[5]; //日付

    const gaichuuNoteName = "【外注ノート】"+" "+gatCliantName+" "+gatJobName+" "+gatDate+" "+"J-id:"+gatJobId+" "+"S-id:"+gatscheduleId
    Logger.log(gaichuuNoteName);



    file.setName(gaichuuNoteName)



    // const befor_text = '案'
    // const after_text = '作業'
    // const regexp = new RegExp( '(?<=' + befor_text + ').*?(?=' + after_text + ')' )
    // const match_text = text.match(regexp)
    // console.log(match_text)
    //console.log(match_text[0])
    //console.log(match_text[1])
    //console.log(match_text[2])
    // console.log(match_text[3])


  }
}












//★★★★★★作業報告書作成★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function createReportVer4(){

  //残物列に記入がないと報告書作成出来ない様にしてる。
  const Leftovers = jobSheet.getRange(numJobSheetCurrentRow,numLeftoversCol+1,1,1).getValue();
  Logger.log(Leftovers);


  if(Leftovers){

    //★★必要な値を取得★★

    const objJob = new JobClass(aryJobSheet[numJobSheetCurrentRow-1]);


    // const gyoumu = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numGyoumuPersonCol];  //業務担当者名を取得
    // const eigyou = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numEigyouPersonCol];  //営業担当者名を取得
    // const date = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numFirstShippingDateCol];  //出荷日を取得
    // const cliant =  rngJobSheet.getValues()[numJobSheetCurrentRow-1][numCliantNameCol];  //顧客名を取得
    // const jobName = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numJobNameCol];  //案件名を取得
    // const cliantPerson =  rngJobSheet.getValues()[numJobSheetCurrentRow-1][numCliantPersonCol];  //顧客担当者を取得
    // const jobNumber = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numJobNumberCol];  //受注番号を取得
    // const jobId = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numJobIdCol];  //案件IDを取得
    // const jobRemark = rngJobSheet.getValues()[numJobSheetCurrentRow-1][numJobRemarkCol];  //備考を取得


    function setValuesTemplateReportFile(){
      ssReport.getRange("E7").setValue(objJob.jobId);
      ssReport.getRange("M3").setValue(objJob.jobGyoumuPerson);
      ssReport.getRange("M4").setValue(objJob.jobEigyouPerson);
      ssReport.getRange("E3").setValue(objJob.jobCliant);
      ssReport.getRange("E5").setValue(objJob.jobName);
      ssReport.getRange("H2").setValue(objJob.jobFirstShippingDate);
      //ssReport.getRange("D7").setValue(quantity);
      //ssReport.getRange("M4").setValue(method);
      ssReport.getRange("E4").setValue(objJob.jobCliantPerson);
      ssReport.getRange("E6").setValue(objJob.jobNum);
      ssReport.getRange("T15").setValue(objJob.jobRemark);

      //↓↓↓↓↓出荷日を構造化データで取得↓↓↓↓
      var formatDate = Utilities.formatDate(ssReport.getRange("H2").getValue(),"Asia/Tokyo","yyyy/MM/dd");

      //↓↓↓↓↓作業報告書に名前を付ける↓↓↓↓    
      var reportName = "【作業報告書】　"+
      formatDate +" "+ 
      ssReport.getRange("E3").getValue() +"様 "+ 
      "【"+ssReport.getRange("E5").getValue()+"】"+
      "J-"+ssReport.getRange("E7").getValue()

      ssReport.rename(reportName)

    }


    const jobFolderId = jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).getValue();
    Logger.log(jobFolderId);


    //★★作業報告書テンプレートを担当者フォルダーにコピー作成★★
    switch (objJob.jobGyoumuPerson){
      case "松澤":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const matsuzawaJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyMatsuzawa = templateReportFile.makeCopy("松澤　作業報告書", matsuzawaJobFolder);
          const templateReportCopyMatsuzawaId = templateReportCopyMatsuzawa.getId(); 
          const templateReportCopyMatsuzawaUrl = templateReportCopyMatsuzawa.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyMatsuzawaId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyMatsuzawaUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;
      
      case "薮田":
        if (jobFolderId){
          const yabutaJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyYabuta = templateReportFile.makeCopy("薮田　作業報告書", yabutaJobFolder);
          const templateReportCopyYabutaId = templateReportCopyYabuta.getId(); 
          const templateReportCopyYabutaUrl = templateReportCopyYabuta.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyYabutaId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyYabutaUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "八子":
        if (jobFolderId){
          const yakoJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyYako = templateReportFile.makeCopy("八子　作業報告書", yakoJobFolder);
          const templateReportCopyYakoId = templateReportCopyYako.getId(); 
          const templateReportCopyYakoUrl = templateReportCopyYako.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyYakoId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyYakoUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "めぐみ":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const uchiyamaJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyUchiyama = templateReportFile.makeCopy("内山　作業報告書", uchiyamaJobFolder);
          const templateReportCopyUchiyamaId = templateReportCopyUchiyama.getId(); 
          const templateReportCopyUchiyamaUrl = templateReportCopyUchiyama.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyUchiyamaId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyUchiyamaUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "北本":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const kitamotoJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyKitamoto = templateReportFile.makeCopy("北本　作業報告書", kitamotoJobFolder);
          const templateReportCopyKitamotoId = templateReportCopyKitamoto.getId(); 
          const templateReportCopyKitamotoUrl = templateReportCopyKitamoto.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyKitamotoId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyKitamotoUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "黄木":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const ohkiJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyOhki = templateReportFile.makeCopy("黄木　作業報告書", ohkiJobFolder);
          const templateReportCopyOhkiId = templateReportCopyOhki.getId(); 
          const templateReportCopyOhkiUrl = templateReportCopyOhki.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyOhkiId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyOhkiUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "橋本":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const hashimotoJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyHashimoto = templateReportFile.makeCopy("橋本　作業報告書", hashimotoJobFolder);
          const templateReportCopyHashimotoId = templateReportCopyHashimoto.getId(); 
          const templateReportCopyHashimotoUrl = templateReportCopyHashimoto.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyHashimotoId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyHashimotoaUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "岩本":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const iwamotoJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyIwamoto = templateReportFile.makeCopy("岩本　作業報告書", iwamotoJobFolder);
          const templateReportCopyIwamotoaId = templateReportCopyIwamoto.getId(); 
          const templateReportCopyIwamotoUrl = templateReportCopyMIwamoto.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyIwamotoId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyIwamotoUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "柿添":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const kakizoeJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyKakizoe = templateReportFile.makeCopy("柿添　作業報告書", kakizoeJobFolder);
          const templateReportCopyKakizoeId = templateReportCopyKakizoe.getId(); 
          const templateReportCopyKakizoeUrl = templateReportCopyKakizoe.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyKakizoeId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyKakizoeUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;

      case "山田":
        if (jobFolderId){
          Logger.log("既に案件フォルダーが有ります")
          const yamadaJobFolder = DriveApp.getFolderById(jobFolderId);
          const templateReportCopyYamada = templateReportFile.makeCopy("山田　作業報告書", yamadaJobFolder);
          const templateReportCopyYamadaId = templateReportCopyYamada.getId(); 
          const templateReportCopyYamadaUrl = templateReportCopyYamada.getUrl();
          var ssReport = SpreadsheetApp.openById(templateReportCopyYamadaId);
          setValuesTemplateReportFile();
          jobSheet.getRange(numJobSheetCurrentRow,numJobReportCol+1,1,1).setValue(templateReportCopyYamadaUrl); 
        }else{
          Logger.log("先に案件フォルダーを作ってください")
          Browser.msgBox("先に案件フォルダーを作ってください");
        };
        break;
      }
    
    }else{
      Browser.msgBox("残物セルに入力してください")
  }
}









//★★★★★★予定表をソートし罫線を描画★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function createBordar(){

  scheduleSheet.getRange(2,1,numScheduleSheetLastRow-1,numScheduleSheetLastCol).sort(
    [{column:numScheduleShippingDateCol+1, ascending: true},
    {column: numShippngMethodCol+1, ascending: true}]
  );

  const aryScheduleValues = scheduleSheet.getDataRange().getValues();

  scheduleSheet.getRange(1,1,scheduleSheet.getMaxRows(),21).setBorder(false,false,false,false,false,false);  //罫線を一度クリア

  for (var i=1; i<numScheduleSheetLastRow-1;i++){
      var dateobject1 = aryScheduleValues[i][12]
      var dateobject2 = aryScheduleValues[i+1][12]

    if (Date.prototype.isPrototypeOf(dateobject2)){
      var date1 = aryScheduleValues[i][12].getDate();
      var date2 = aryScheduleValues[i+1][12].getDate();
      if (date1!=date2){
        Logger.log("date1とdate2は違います")
      //  scheduleSheet.getRange(i+1,1,1,21).setBorder(null,null,true,null,null,null, 'blue' ,null,);
        scheduleSheet.getRange(i+1,1,1,21).setBorder(null,null,true,null,null,null, 'blue',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);  
      }else{
        Logger.log("date1とdate2は同じです")
      }
    }else{
      Logger.log("出荷日に日付以外の値が入力されてます("+(i+2)+")行目")
      Browser.msgBox("出荷日が未入力か日付以外の値が入力されてます。　【"+(i+2)+"】行目")
      break
    }
  }  
  
}









//◆◆◆↓↓↓↓↓↓以下 jobId に　var　を使ってる。　修正必要？！！　↓↓↓↓↓↓◆◆◆

//★★★★★★案件登録後、予定に入力★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
function copyValueFromJobSheetToSchduleSheet(){

  createJobFolder()

　　if (ss.getActiveSheet().getName()!="予定"){


　　//if (アクティブユーザー＝＝カレントロー担当者){}  アクティブユーザーを確認し合っていたら予定表を作成　
      const activeuserId = Session.getActiveUser(). getUserLoginId()
      Logger.log(activeuserId);
      

    //★★新規案件登録の場合と既存の案件の場合でjobIdの取得と予定表への書き込みに場合分けしてる。★★
    if (jobSheet.getRange(numJobSheetCurrentRow,2,1,1).isBlank()){
      var jobId = jobSheet.getRange(numJobSheetCurrentRow,2,1,1).setFormula("=DEC2HEX(RANDBETWEEN(0,4294967295),8)").getValue();
      jobSheet.getRange(numJobSheetCurrentRow,2,1,1).setValue(jobId);
      Logger.log(numJobSheetCurrentRow);
      Logger.log(jobId);
      Logger.log("新たにUniqueId生成、予定表に記載");
    }else{
      var jobId = jobSheet.getRange(numJobSheetCurrentRow,2,1,1).getValue();
      jobSheet.getRange(numJobSheetCurrentRow,2,1,1).setValue(jobId);
      Logger.log(numJobSheetCurrentRow);
      Logger.log(jobId);
      Logger.log("生成ずみUniqueIdを使い、予定表に記載");
    }
    
    //案件登録シートより値を取得
    Logger.log(numJobSheetCurrentRow);
    Logger.log(numJobNumberCol);

    const jobNum = aryJobSheet[numJobSheetCurrentRow-1][numJobNumberCol];
//    Logger.log("aaaaa"+jobNum[numJobSheetCurrentRow-1][numJobNumberCol]);
    const cliant = aryJobSheet[numJobSheetCurrentRow-1][numCliantNameCol];
    const cliantPerson = aryJobSheet[numJobSheetCurrentRow-1][numCliantPersonCol];
    const jobName = aryJobSheet[numJobSheetCurrentRow-1][numJobNameCol];
    const jobDetail = aryJobSheet[numJobSheetCurrentRow-1][numJobDetailCol];
    const gyoumuPerson =aryJobSheet[numJobSheetCurrentRow-1][numGyoumuPersonCol];
    const eigyouPerson = aryJobSheet[numJobSheetCurrentRow-1][numEigyouPersonCol];
    const firstShippingDate = aryJobSheet[numJobSheetCurrentRow-1][numFirstShippingDateCol];
    const jobFolderId = jobSheet.getRange(numJobSheetCurrentRow,numJobFolderIdCol+1,1,1).getValue();
    //const jobFolderId = aryJobSheet[numJobSheetCurrentRow-1][numJobFolderIdCol];

    //案件シート初出荷日より値を取得し年月に変換し該当カラムに入力
    const rngJobSheetYearMonth = jobSheet.getRange(numJobSheetCurrentRow,numJobYearMonth,1,1);
    const jobSheetYear = jobSheet.getRange(numJobSheetCurrentRow,numFirstShippingDateCol+1,1,1).getValue().getFullYear();
    const jobSheetMonth = jobSheet.getRange(numJobSheetCurrentRow,numFirstShippingDateCol+1,1,1).getValue().getMonth()+1;
    Logger.log(jobSheetYear +"-"+ jobSheetMonth);
    const jobSheetYearMonth = jobSheetYear +"-"+ jobSheetMonth;
    rngJobSheetYearMonth.setValue(jobSheetYearMonth);


    //予定シートに値を入力
    const scheduleId = scheduleSheet.getRange(numScheduleSheetLastRow+1,2,1,1).setFormula("=DEC2HEX(RANDBETWEEN(0,4294967295),8)").getValue();
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleStatusCol+1,1,1).setValue("未着手");
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleIdCol+1,1,1).setValue(scheduleId);  
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleJobIdCol+1,1,1).setValue(jobId);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleJobNumberCol+1,1,1).setValue(jobNum);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleCliantNameCol+1,1,1).setValue(cliant);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleCliantPersonCol+1,1,1).setValue(cliantPerson);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleJobNameCol+1,1,1).setValue(jobName);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleJobDetailCol+1,1,1).setValue(jobDetail);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleGyoumuPersonCol+1,1,1).setValue(gyoumuPerson);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleEigyouPersonCol+1,1,1).setValue(eigyouPerson);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleShippingDateCol+1,1,1).setValue(firstShippingDate);
    scheduleSheet.getRange(numScheduleSheetLastRow+1,numScheduleJobFolderIdCol+1,1,1).setValue(jobFolderId);

    //作業予定表にアクティブセルを移動する関数を呼び出してる。
  　//setActiveSheet();
    SpreadsheetApp.setActiveSheet(scheduleSheet);
    const range = SpreadsheetApp.getActiveSheet().getRange(numScheduleSheetLastRow+1,8,1,1);
    SpreadsheetApp.setActiveRange(range);
  }else{
    Browser.msgBox("案件登録シートから予定を入力してください");
  }
}
