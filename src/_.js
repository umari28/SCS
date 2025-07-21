function onOpen() {

    var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
    var menu = ui.createMenu('コマンド');  // Uiクラスからメニューを作成する
    menu.addItem('インフォ', 'interInfo');
    menu.addItem('シャッフル', 'shuffle');
    menu.addItem('OPEN','do_open');
    menu.addItem('新規クラス作成','make_new_class');
    menu.addItem('ハードリセット','hard_reset');
    menu.addItem('クラス呼び出し','call_class');
    menu.addToUi();                            // メニューをUiクラスに追加する
}



const ss = SpreadsheetApp.getActiveSpreadsheet();
const dec_sheet = ss.getSheetByName("座席指定");
const dev_sheet = ss.getSheetByName("管理者用シート");
const all_class = ss.getSheetByName("クラス情報一覧");
const name_sheet = ss.getSheetByName("名前");
const main_sheet = ss.getSheetByName("メインシート");
const br_sheet = ss.getSheetByName("暗転");
const br2_sheet = ss.getSheetByName("暗転２");
const open_sheet = ss.getSheetByName("明転");
const row_array = ["B","E","H","K","N","Q"];

function interInfo(){
    const maxRows = all_class.getMaxRows();
    const lastRow = all_class.getRange(maxRows, 2).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const nowlastRow = return_LastRow(name_sheet, 3);
    // var nowlastRow = name_sheet.getRange()
    console.log(nowlastRow);

    main_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    br_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    br2_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    dev_sheet.getRange(1,2,42,1).clearContent();

    let count = 1;
    let now_Row = 0;

    for(let i = 2; i <= 17; i = i + 3){
        for(let j = 7; j <= 25; j = j + 3){
            if(dec_sheet.getRange(j,i).getValue() == true){
                dev_sheet.getRange(count,2).setValue(count);
                main_sheet.getRange(j,i).setValue("='管理者用シート'!B" + count);
                main_sheet.getRange(j , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'名前\'!A1:C' + nowlastRow + ',2)');
                main_sheet.getRange(j + 1 , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'名前\'!A1:C' + nowlastRow + ',3)');

                open_sheet.getRange(j,i).setValue("='管理者用シート'!B" + count);
                open_sheet.getRange(j , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'名前\'!A1:C' + nowlastRow + ',2)');
                open_sheet.getRange(j + 1 , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'名前\'!A1:C' + nowlastRow + ',3)');

                setColor(br_sheet, "gray", j, i);
                setColor(br_sheet, "gray", j, i+1);
                setColor(br_sheet, "gray", j+1, i+1);

                if(dec_sheet.getRange(j,i).getBackground() != '#ffffff'){
                    setColor(br2_sheet, "gray", j, i);
                    setColor(br2_sheet, "gray", j, i+1);
                    setColor(br2_sheet, "gray", j+1, i+1);
                }else{
                    br2_sheet.getRange(j,i).setValue("='名前'!E" + count);
                }

                count++;
            }
        }
        now_Row++;
    }
    // all_class.getRange(lastRow + 1, 1).setValue(main_sheet.getRange(15,21).getValue());
    // all_class.getRange(lastRow + 1, 2).setValue(count - 1); //名前のシートにクラスの人数を記載

}

function shuffle(){
    let row = return_LastRow(all_class, 1) + 1;
    for(let i = 2; i <= row; i++){
        if(main_sheet.getRange(15,21).getValue() == all_class.getRange(i,1).getValue()){
            var laRo = all_class.getRange(i,2).getValue();
            break;
        }
    }

    for(i = 0; i<10; i++){

      let range = dev_sheet.getRange(1,2,laRo,1);
      range.randomize();

    }

    let dec_array = dev_sheet.getRange(1,2,laRo,1).getValues();
    console.log(dec_array);

    br_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
    main_sheet.getRange(25,21).setValue("0");

}

function do_open(){
    var nowNum = main_sheet.getRange(25,21).getValue();
    if(nowNum == "0"){
        br2_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
        main_sheet.getRange(25,21).setValue("1");
    }else if(nowNum == "1"){
        open_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
    }
}

function hard_reset(){

    var check = Browser.msgBox("本当にリセットしますか","バックアップを取る事をおすすめします",Browser.Buttons.OK_CANCEL);

    if(check = "ok"){
        main_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        main_sheet.getRange(25,21).setValue("0");
        br_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        br2_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        let laRo = return_LastRow(dev_sheet, 2);
        dev_sheet.getRange(1,2,laRo,1).clearContent();
        let laRoo = return_LastRow(all_class, 1);
        all_class.getRange(2,1,laRoo,2).clearContent();
    }

}

function make_new_class(){
    var check = Browser.msgBox("メインシートに配置した席順で新規クラスが作成されます。\n新規クラス作成前にデータセットしてください。",Browser.Buttons.OK_CANCEL);
    if(check == "cancel"){
        return 0;
    }


    var newName = Browser.inputBox("新しいクラス名を入力してください");
    let sheet_num = ss.getNumSheets(); //シートの数

    let laRo = return_LastRow(all_class, 1); //1
    for(let i = 2; i <= laRo + 1 ; i++){
        if(all_class.getRange(i,1).getValue() == newName){
            Browser.msgBox("同じクラス名で作成することはできません。");
            return 0;
        }
    }
    all_class.getRange(laRo + 1, 1).setValue(newName);
    let name_Row = return_LastRow(name_sheet, 2);
    all_class.getRange(laRo + 1, 2).setValue(name_sheet.getRange(name_Row, 1).getValue());


    var newSheets = ss.insertSheet().setName(newName); //新しいシートを追加する

  //席のレイアウトコピー
  main_sheet.getRange(1,1,27,19).copyTo(newSheets.getRange(1,1));
  newSheets.setHiddenGridlines(true);
  var newSheets = ss.getSheetByName(newName);


  // 列、行の幅を変更
  for(i = 7; i <= 25; i = i +3){

    newSheets.setRowHeights(i, i + 1, 30);
    newSheets.setRowHeight(i + 2, 21);

  }

  for(i = 1; i <= 19; i++){

    var str = i % 3;
    console.log(str);

    switch(str){

      case 1:
        newSheets.setColumnWidth(i,20);
        break;
      case 2:
        newSheets.setColumnWidth(i,50);
        break;
      case 0:
        newSheets.setColumnWidth(i,168);
        break;

    }
  }

  const pullList = all_class.getRange(1, 1, laRo + 1, 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(pullList).build();
  const cell = main_sheet.getRange(15,21);
  cell.setDataValidation(rule);
}

function call_class(){
    let call_class = main_sheet.getRange(15,21).getValue();
    call_class.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
}


function setColor(sheet, color, i, j){
    if(color == "gray"){
        sheet.getRange(i, j).setBackground('#999999');
    }
}

function return_LastRow(sheet, row){
    const max_Row = sheet.getMaxRows();
    return sheet.getRange(max_Row, row).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

function make_nameList(){
    let laRo = return_LastRow(nameList, 2);
    var nameList = name_sheet.getRange(1,2,laRo,2);
    return nameList;
}

function debug(){
    var nowlastRow = return_LastRow(name_sheet, 3);
    // var nowlastRow = name_sheet.getRange(1,3).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    console.log(nowlastRow);
}

// 配列データをスプレッドシートに保存する関数
function saveArrayToSheet(arrayData, sheetName, startRow = 1, startCol = 1) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        // シートが存在しない場合は新規作成
        const newSheet = ss.insertSheet(sheetName);
        newSheet.getRange(startRow, startCol, arrayData.length, arrayData[0].length).setValues(arrayData);
    } else {
        // 既存シートに上書き
        sheet.getRange(startRow, startCol, arrayData.length, arrayData[0].length).setValues(arrayData);
    }
}

// スプレッドシートから配列データを読み込む関数
function loadArrayFromSheet(sheetName, startRow = 1, startCol = 1, numRows = null, numCols = null) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        console.log(`シート "${sheetName}" が見つかりません`);
        return null;
    }
    
    if (numRows === null) {
        numRows = sheet.getLastRow() - startRow + 1;
    }
    if (numCols === null) {
        numCols = sheet.getLastColumn() - startCol + 1;
    }
    
    return sheet.getRange(startRow, startCol, numRows, numCols).getValues();
}

// 1次元配列を保存する関数
function saveSimpleArray(arrayData, sheetName, startRow = 1, startCol = 1) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        const newSheet = ss.insertSheet(sheetName);
        newSheet.getRange(startRow, startCol, arrayData.length, 1).setValues(arrayData.map(item => [item]));
    } else {
        sheet.getRange(startRow, startCol, arrayData.length, 1).setValues(arrayData.map(item => [item]));
    }
}

// 1次元配列を読み込む関数
function loadSimpleArray(sheetName, startRow = 1, startCol = 1, numRows = null) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        console.log(`シート "${sheetName}" が見つかりません`);
        return null;
    }
    
    if (numRows === null) {
        numRows = sheet.getLastRow() - startRow + 1;
    }
    
    const values = sheet.getRange(startRow, startCol, numRows, 1).getValues();
    return values.map(row => row[0]); // 1次元配列に変換
}

// PropertiesServiceを使用して配列データを保存する関数
function saveArrayToProperties(arrayData, key) {
    const properties = PropertiesService.getScriptProperties();
    const jsonString = JSON.stringify(arrayData);
    properties.setProperty(key, jsonString);
}

// PropertiesServiceから配列データを読み込む関数
function loadArrayFromProperties(key) {
    const properties = PropertiesService.getScriptProperties();
    const jsonString = properties.getProperty(key);
    if (jsonString) {
        return JSON.parse(jsonString);
    }
    return null;
}

// 複数の配列を一括で保存する関数
function saveMultipleArrays(arraysObject) {
    const properties = PropertiesService.getScriptProperties();
    const propertiesToSet = {};
    
    for (const [key, value] of Object.entries(arraysObject)) {
        propertiesToSet[key] = JSON.stringify(value);
    }
    
    properties.setProperties(propertiesToSet);
}

// 複数の配列を一括で読み込む関数
function loadMultipleArrays(keys) {
    const properties = PropertiesService.getScriptProperties();
    const result = {};
    
    for (const key of keys) {
        const jsonString = properties.getProperty(key);
        if (jsonString) {
            result[key] = JSON.parse(jsonString);
        }
    }
    
    return result;
}

// 使用例：配列データの永続化テスト
function testArrayPersistence() {
    // テスト用の配列データ
    const testArray = [1, 2, 3, 4, 5];
    const testMatrix = [
        ['A', 'B', 'C'],
        [1, 2, 3],
        ['X', 'Y', 'Z']
    ];
    
    // 方法1: スプレッドシートに保存
    saveArrayToSheet(testMatrix, 'テストデータ', 1, 1);
    saveSimpleArray(testArray, 'テスト配列', 1, 1);
    
    // 方法2: PropertiesServiceに保存
    saveArrayToProperties(testArray, 'testArray');
    saveArrayToProperties(testMatrix, 'testMatrix');
    
    // 複数配列の一括保存
    saveMultipleArrays({
        'userList': ['田中', '佐藤', '鈴木'],
        'scores': [85, 92, 78],
        'settings': {theme: 'dark', language: 'ja'}
    });
    
    console.log('データを保存しました');
}

// 使用例：保存したデータの読み込み
function loadPersistedData() {
    // スプレッドシートから読み込み
    const loadedMatrix = loadArrayFromSheet('テストデータ');
    const loadedArray = loadSimpleArray('テスト配列');
    
    // PropertiesServiceから読み込み
    const propsArray = loadArrayFromProperties('testArray');
    const propsMatrix = loadArrayFromProperties('testMatrix');
    
    // 複数配列の一括読み込み
    const multipleData = loadMultipleArrays(['userList', 'scores', 'settings']);
    
    console.log('読み込んだデータ:', {
        loadedMatrix,
        loadedArray,
        propsArray,
        propsMatrix,
        multipleData
    });
    
    return {
        loadedMatrix,
        loadedArray,
        propsArray,
        propsMatrix,
        multipleData
    };
}