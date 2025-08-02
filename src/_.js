function onOpen() {

    var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
    var menu = ui.createMenu('コマンド');  // Uiクラスからメニューを作成する
    menu.addItem('データセット', 'interInfo');
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
const name_sheet = ss.getSheetByName("名簿");
const add_name = ss.getSheetByName("クラス追加用名簿");
const main_sheet = ss.getSheetByName("メインシート");
const br_sheet = ss.getSheetByName("暗転");
const br2_sheet = ss.getSheetByName("暗転２");
const open_sheet = ss.getSheetByName("明転");
const temp_sheet = ss.getSheetByName("はくシート");
const row_array = ["B","E","H","K","N","Q"];
const alpha = ["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

//FUNCTIONS:   情報入力
function interInfo(){
    const maxRows = all_class.getMaxRows();
    const lastRow = all_class.getRange(maxRows, 2).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const nowlastRow = return_LastRow(dev_sheet, 6);
    // var nowlastRow = name_sheet.getRange()
    console.log(nowlastRow);

    main_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    br_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    br2_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
    dev_sheet.getRange(1,2,42,1).clearContent();

    var count = 0;
    for(let i = 2; i <= 17; i = i + 3){
        for(let j = 7; j <= 25; j = j + 3){
            if(dec_sheet.getRange(j,i).getValue() == true){
                count++;
            }
        }
    }

    if(return_LastRow(add_name, 2) != count){
        Browser.msgBox("名簿の人数と席の数が一致しません。\n名簿の人数を確認してください。");
        return 0;
    }

    var count = 1;
    let now_Row = 0;

    for(let i = 2; i <= 17; i = i + 3){
        for(let j = 7; j <= 25; j = j + 3){
            if(dec_sheet.getRange(j,i).getValue() == true){
                dev_sheet.getRange(count,2).setValue(count);
                main_sheet.getRange(j,i).setValue("='管理者用シート'!B" + count);
                main_sheet.getRange(j , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'管理者用シート\'!E1:G' + nowlastRow + ',2)');
                main_sheet.getRange(j + 1 , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'管理者用シート\'!E1:G' + nowlastRow + ',3)');

                open_sheet.getRange(j,i).setValue("='管理者用シート'!B" + count);
                open_sheet.getRange(j , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'管理者用シート\'!E1:G' + nowlastRow + ',2)');
                open_sheet.getRange(j + 1 , i + 1 ).setValue('=VLOOKUP(' + row_array[now_Row] + j + ',\'管理者用シート\'!E1:G' + nowlastRow + ',3)');

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

//FUNCTIONS:   シャッフル
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

    let count = 0;
    let now_class = main_sheet.getRange(15,21).getValue();

    for(let i = 2; i <= 17; i = i + 3){
        for(let j = 7; j <= 25; j = j + 3){
            now_class.getRange(j,i).setValue(dec_array[count]);
            count++;
        }
    }

    br_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
    main_sheet.getRange(25,21).setValue("0");

}

//FUNCTIONS:   明転
function do_open(){
    var nowNum = main_sheet.getRange(25,21).getValue();
    if(nowNum == "0"){
        br2_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
        main_sheet.getRange(25,21).setValue("1");
    }else if(nowNum == "1"){
        open_sheet.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
    }
}

//FUNCTIONS:  ハードリセット
function hard_reset(){

    var check = Browser.msgBox("本当にリセットしますか","バックアップを取る事をおすすめします",Browser.Buttons.OK_CANCEL);

    if(check == "ok"){
        main_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        main_sheet.getRange(25,21).setValue("0");
        br_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        br2_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
        let laRo = return_LastRow(dev_sheet, 2);
        dev_sheet.getRange(1,2,laRo,1).clearContent();
        let laRoo = return_LastRow(all_class, 1);
        all_class.getRange(2,1,laRoo,2).clearContent();
    }else{
        Browser.msgBox("リセットをキャンセルしました");
    }

}

//FUNCTIONS:   新規クラス作成
function make_new_class(){
    var check = Browser.msgBox("一度データセットしてください。",Browser.Buttons.OK_CANCEL);
    if(check == "cancel"){
        return 0;
    }

    dev_sheet.getRange(1,2,42,1).clearContent();

    var newName = Browser.inputBox("新しいクラス名を入力してください");
    let sheet_num = ss.getNumSheets(); //シートの数

    var classMember_column = Browser.inputBox("このクラスの名簿の一番左の列のアルファベットを入力してください");

    var member_quantity = Browser.inputBox("クラスの人数を入力してください");

    let laRo = return_LastRow(all_class, 1); //1
    for(let i = 2; i <= laRo + 1 ; i++){
        if(all_class.getRange(i,1).getValue() == newName){
            Browser.msgBox("同じクラス名で作成することはできません。");
            return 0;
        }
    }

    all_class.getRange(laRo + 1, 1).setValue(newName);
    // let name_Row = return_LastRow(name_sheet, 2);
    // all_class.getRange(laRo + 1, 2).setValue(name_sheet.getRange(name_Row, 1).getValue());
    all_class.getRange(laRo + 1, 2).setValue(member_quantity);


    // var newSheets = ss.insertSheet().setName(newName); //新しいシートを追加する
    var newSheets = temp_sheet.copyTo(ss).setName(newName);

  //席のレイアウトコピー
//   main_sheet.getRange(1,1,27,19).copyTo(newSheets.getRange(1,1));
//   newSheets.setHiddenGridlines(true);
//   var newSheets = ss.getSheetByName(newName);


//   // 列、行の幅を変更
//   for(i = 7; i <= 25; i = i +3){

//     newSheets.setRowHeights(i, i + 1, 30);
//     newSheets.setRowHeight(i + 2, 21);

//   }

//   for(i = 1; i <= 19; i++){

//     var str = i % 3;
//     console.log(str);

//     switch(str){

//       case 1:
//         newSheets.setColumnWidth(i,20);
//         break;
//       case 2:
//         newSheets.setColumnWidth(i,50);
//         break;
//       case 0:
//         newSheets.setColumnWidth(i,168);
//         break;

//     }
//   }

  //メインシートにプルダウンを作成
  const pullList = all_class.getRange(1, 1, laRo + 1, 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(pullList).build();
  const cell = main_sheet.getRange(15,21);
  cell.setDataValidation(rule);

  //名簿の下にクラス名を入力
  let col = comvert_alphabet_num(classMember_column);
  name_sheet.getRange(43,col).setValue(newName);

//   main_sheet.getRange(7,2,20,17).clearContent().setBackground('#ffffff');
}

//TODO:   クラス削除
function delete_class(){
    // var check = Browser.msgBox("本当に削除しますか","バックアップを取る事をおすすめします",Browser.Buttons.OK_CANCEL);
    let name = Browser.inputBox("削除するクラス名を入力してください");
    ss.deleteSheet(ss.getSheetByName(name));
    for(let i = 1; i <= return_LastColumn(name_sheet, 43); i = i + 3){
        if(name_sheet.getRange(43,i).getValue() == name){
            name_sheet.deleteColumn(i,i+2);
            name_sheet.insertColumnsAfter(return_LastColumn(name_sheet, 43),3);
            break;
        }
    }
}

//FUNCTIONS:   クラス呼び出し
function call_class(){
    let call_class = main_sheet.getRange(15,21).getValue();
    // call_class.getRange(1,1,27,19).copyTo(main_sheet.getRange(1,1));
    for(let i = 1; i <= return_LastColumn(name_sheet, 43); i = i + 3){
        if(name_sheet.getRange(43,i).getValue() == call_class){
            name_sheet.getRange(1,i+1,return_LastRow(name_sheet, i+1),2).copyTo(dev_sheet.getRange(1,2));
        }
    }
}

//CLASS:
function setColor(sheet, color, i, j){
    if(color == "gray"){
        sheet.getRange(i, j).setBackground('#999999');
    }
}

//CLASS:
function return_LastRow(sheet, row){
    const max_Row = sheet.getMaxRows();
    return sheet.getRange(max_Row, row).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

//CLASS:
function return_LastColumn(sheet, column){
    const max_Column = sheet.getMaxColumns();
    console.log(max_Column);
    return sheet.getRange(column,max_Column).getNextDataCell(SpreadsheetApp.Direction.LEFT).getColumn();
    //const lastColumn = sheet.getRange(1, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
}

//FUNCTIONS:
function make_nameList(){
    let laRo = return_LastRow(nameList, 2);
    var nameList = name_sheet.getRange(1,2,laRo,2);
    return nameList;
}

//FUNCTIONS:
function debug(){
    var nowlastRow = return_LastRow(name_sheet, 3);
    // var nowlastRow = name_sheet.getRange(1,3).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    console.log(nowlastRow);
}

//CLASS:    アルファベット数字変換
function comvert_alphabet_num(alpha){
    switch(alpha){
        case "A":
            return 1;
        case "B":
            return 2;
        case "C":
            return 3;
        case "D":
            return 4;
        case "E":
            return 5;
        case "F":
            return 6;
        case "G":
            return 7;
        case "H":
            return 8;
        case "I":
            return 9;
        case "J":
            return 10;
        case "K":
            return 11;
        case "L":
            return 12;
        case "M":
            return 13;
        case "N":
            return 14;
        case "O":
            return 15;
        case "P":
            return 16;
        case "Q":
            return 17;
        case "R":
            return 18;
        case "S":
            return 19;
        case "T":
            return 20;
        case "U":
            return 21;
        case "V":
            return 22;
        case "W":
            return 23; 
        case "X":
            return 24;
        case "Y":
            return 25;
        case "Z":
            return 26;
    }
}