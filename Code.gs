function onOpen(e) {
  var ui = SlidesApp.getUi()
  var menu = ui.createAddonMenu() // メニューの追加、（セルを結合した表が一つでもあると、処理が行われません。）
  menu
  .addItem('[全スライド]DeepL翻訳(英>日)', 'translateAllDeepL')
  .addItem('[特定ページ]DeepL翻訳(英>日)','specifiedPageTranslateDeepL')
  .addItem('[全スライド]Google翻訳(英>日)', 'translateAll')
  .addItem('[特定ページ]Google翻訳(英>日)','specifiedPageTranslate')
  .addToUi();  
}

function buildSearchCard(opt_error) {
  var banner = CardService.newImage()
  .setImageUrl('https://2.bp.blogspot.com/-8-ypfalviWA/WjCryBf7m9I/AAAAAAABIx4/3VAOGejdWogMnchjAO_F2rJG-ZeOmWJEQCLcBGAs/s400/building_nihongo_gakkou.png');

  var ontranslateAllDeepLAction = CardService.newAction()
  .setFunctionName("translateAllDeepL")
  .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var translateAllDeepLButton = CardService.newTextButton()
  .setText("[全スライド]DeepL翻訳(英>日)")
  .setOnClickAction(ontranslateAllDeepLAction);

  var onspecifiedPageTranslateDeepLAction = CardService.newAction()
  .setFunctionName("specifiedPageTranslateDeepL")
  .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var specifiedPageTranslateDeepLButton = CardService.newTextButton()
  .setText("[特定ページ]DeepL翻訳(英>日)")
  .setOnClickAction(onspecifiedPageTranslateDeepLAction);

  var ontranslateAllAction = CardService.newAction()
  .setFunctionName("translateAll")
  .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var translateAllButton = CardService.newTextButton()
  .setText("[全スライド]Google翻訳(英>日)")
  .setOnClickAction(ontranslateAllAction);

  var onspecifiedPageTranslateAction = CardService.newAction()
  .setFunctionName("specifiedPageTranslate")
  .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var specifiedPageTranslateButton = CardService.newTextButton()
  .setText("[特定ページ]Google翻訳(英>日)")
  .setOnClickAction(onspecifiedPageTranslateAction);

  var section = CardService.newCardSection()
  .addWidget(banner)
  .addWidget(translateAllDeepLButton)
  .addWidget(specifiedPageTranslateDeepLButton)
  .addWidget(translateAllButton)
  .addWidget(specifiedPageTranslateButton);

  if (opt_error) {
    var message = CardService.newTextParagraph()
    .setText("Note: " + opt_error);
    section.addWidget(message);
  }

  return CardService.newCardBuilder()
  .addSection(section)
  .build();
}

function onHomePage() {
  var card = buildSearchCard();
  return [card];
}
 
const apiKey = 'YOUR_DEEPL_API_KEY';
const apiUrl = 'https://api-free.deepl.com/v2/translate';

function deepltranslate(text, src, tgt) {
  let t = text.replace( '&', 'and' ).toString();
  let content = encodeURI(`auth_key=${apiKey}&text=${t}&source_lang=${src}&target_lang=${tgt}`);
 
  const postheader = {
    "accept":"gzip, */*",
    "timeout":"20000",
    "Content-Type":"application/x-www-form-urlencoded"
  } 
 
  const parameters = {
    "method": "post",
    "headers": postheader,
    'payload': content
  }
 
  /*try {
    let response = UrlFetchApp.fetch(apiUrl, parameters);
  }
  catch (e) {
    Logger.log(e.toString());return 'DeepL:Exception';
  }*/
  let response = UrlFetchApp.fetch(apiUrl, parameters);
  let response_code = response.getResponseCode().toString();
  if (response_code != 200) return `DeepL:HTTP Error(${response_code})`
 
  let json = JSON.parse(response.getContentText('UTF-8'));// JSONからテキストを取り出す
  return json.translations[0].text;
}
 
function translateAllDeepL() {
const presentation  = SlidesApp.getActivePresentation();
const slides        = presentation.getSlides();
 
console.log(presentation.getName());
console.log(slides);
console.log('スライドの枚数: %s',slides.length);
 
for(let i = 0; i < slides.length; i++){
  for(let j = 0; j < slides[i].getPageElements().length; j++){
    /*console.log(slides[i].getPageElements()[j].getPageElementType().toString())*/
    if (slides[i].getPageElements()[j].getPageElementType().toString() == 'SHAPE') {
 
      const contents = slides[i].getPageElements()[j].asShape().getText().asString();
     
      /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
      const notes    = slides[i].getNotesPage();
      notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
      /*翻訳後のテキストを貼り付け*/
      const results  = deepltranslate(contents, 'en', 'ja');
      slides[i].getPageElements()[j].asShape().getText().setText(results);
    }
    else if (slides[i].getPageElements()[j].getPageElementType().toString() == 'GROUP'){
      const elemsInGroup = slides[i].getPageElements()[j].asGroup().getChildren()
      for(const elem of elemsInGroup) {
        /*console.log(elem.getPageElementType().toString());*/
        if (elem.getPageElementType().toString() === 'SHAPE') {
          const contents = elem.asShape().getText().asString();
     
          /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
          const notes    = slides[i].getNotesPage();
          notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
          /*翻訳後のテキストを貼り付け*/
          const results  = deepltranslate(contents, 'en', 'ja');
          elem.asShape().getText().setText(results);
        }
        else if (elem.getPageElementType().toString() === 'GROUP') {
          const elemsInGroup2 = elem.asGroup().getChildren()
          for(const elem2 of elemsInGroup2) {
            console.log(elem2.getPageElementType().toString());
            if (elem2.getPageElementType().toString() === 'SHAPE') {
              const contents = elem2.asShape().getText().asString();
     
              /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
              const notes    = slides[i].getNotesPage();
              notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
              /*翻訳後のテキストを貼り付け*/
              const results  = deepltranslate(contents, 'en', 'ja');
              elem2.asShape().getText().setText(results);
            }
          }
        }
      }
    }
  }//for_j
}//for_i
/*テーブル内も変換出来るようにする*//*全スライドの取得*/
for(let i = 0; i < slides.length; i++){
  for(let j = 0; j < slides[i].getTables().length; j++){
     
    console.log('スライド　%s　ページ目には、表　%sがあります。',i+1,slides[i].getTables());
     
    const rowIndex    = slides[i].getTables()[j].getNumRows();
    const columnIndex = slides[i].getTables()[j].getNumColumns();
     
    console.log('行数：　%s',rowIndex);
    console.log('列数：　%s',columnIndex);
     
    for(let k = 0; k < rowIndex; k++){
      for(let l = 0; l < columnIndex; l++){
         
        /**/
        const innertable  = slides[i].getTables()[j].getCell(k, l).getText().asString();
        const results     = deepltranslate(innertable, 'en', 'ja');
         
        console.log('getCell(%s, %s)　元のテキスト：　%S',k,l,innertable);
        console.log('getCell(%s, %s)　翻訳後のテキスト：　%S',k,l,results);
         
         
        if(innertable === ''){continue}
        slides[i].getTables()[j].getCell(k, l).getText().setText(results);
      }//for_l
    }//for_k
  }//for_j
}//for_i
}
 
function specifiedPageTranslateDeepL() {
const presentation  = SlidesApp.getActivePresentation();
const slides        = presentation.getSlides();
 
/*画面の操作を変化させるための記述*/
const ui            = SlidesApp.getUi();
const response      = ui.prompt(
  '翻訳こんにゃく',
  '翻訳したいスライドのページ番号を入力してください。',
  ui.ButtonSet.OK_CANCEL
);
 
//スライドページは配列で取得されるため
let page = response.getResponseText();
page -= 1;
 
switch(response.getSelectedButton()){
  case ui.Button.OK:
    console.log('ページ番号は、%s です。', page);
    break;
  
  case ui.Button.CANCEL:
    console.log('キャンセルが押されたため、処理を中断しました。');
    break;
     
  case ui.Button.CLOSE:
    console.log('閉じるボタンが押されました。');
}
 
for(let j = 0; j < slides[page].getPageElements().length; j++){
  /*console.log(slides[page].getPageElements()[j].getPageElementType().toString())*/
  if (slides[page].getPageElements()[j].getPageElementType().toString() == 'SHAPE') {
 
    const contents = slides[page].getPageElements()[j].asShape().getText().asString();
   
    /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
    const notes    = slides[page].getNotesPage();
    notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
    /*翻訳後のテキストを貼り付け*/
    const results  = deepltranslate(contents, 'en', 'ja');
    slides[page].getPageElements()[j].asShape().getText().setText(results);
  }
  else if (slides[page].getPageElements()[j].getPageElementType().toString() == 'GROUP'){
    const elemsInGroup = slides[page].getPageElements()[j].asGroup().getChildren()
    for(const elem of elemsInGroup) {
      /*console.log(elem.getPageElementType().toString());*/
      if (elem.getPageElementType().toString() === 'SHAPE') {
        const contents = elem.asShape().getText().asString();
   
        /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
        const notes    = slides[page].getNotesPage();
        notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
        /*翻訳後のテキストを貼り付け*/
        const results  = deepltranslate(contents, 'en', 'ja');
        elem.asShape().getText().setText(results);
      }
      else if (elem.getPageElementType().toString() === 'GROUP') {
        const elemsInGroup2 = elem.asGroup().getChildren()
        for(const elem2 of elemsInGroup2) {
          console.log(elem2.getPageElementType().toString());
          if (elem2.getPageElementType().toString() === 'SHAPE') {
            const contents = elem2.asShape().getText().asString();
   
            /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
            const notes    = slides[page].getNotesPage();
            notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
            /*翻訳後のテキストを貼り付け*/
            const results  = deepltranslate(contents, 'en', 'ja');
            elem2.asShape().getText().setText(results);
          }
        }
      }
    }
  }
}//for_j
 
/*特定のページ番号を取得*/
for(let i = 0; i < slides[page].getTables().length; i++){
   
  const rowIndex    = slides[page].getTables()[i].getNumRows();
  const columnIndex = slides[page].getTables()[i].getNumColumns();
   
  console.log('行数：　%s',rowIndex);
  console.log('列数：　%s',columnIndex);
   
  for(let j = 0; j < rowIndex; j++){
    for(let k = 0; k < columnIndex; k++){
       
      /**/
      const innertable  = slides[page].getTables()[i].getCell(j, k).getText().asString();
      const results     = deepltranslate(innertable, 'en', 'ja');
       
      console.log('getCell(%s, %s)　元のテキスト：　%S',j,k,innertable);
      console.log('getCell(%s, %s)　翻訳後のテキスト：　%S',j,k,results);
       
      if(innertable === ''){continue}
      slides[page].getTables()[i].getCell(j, k).getText().setText(results);
    }//for_k
  }//for_j
}//for_i
}
 
/*スライド内容全体を英語に翻訳する。表や図形（シェイプ）の上に書かれた文言については翻訳されない。*/
function translateAll() {
const presentation  = SlidesApp.getActivePresentation();
const slides        = presentation.getSlides();
 
console.log(presentation.getName());
console.log(slides);
console.log('スライドの枚数: %s',slides.length);
 
for(let i = 0; i < slides.length; i++){
  for(let j = 0; j < slides[i].getPageElements().length; j++){
    /*console.log(slides[i].getPageElements()[j].getPageElementType().toString())*/
    if (slides[i].getPageElements()[j].getPageElementType().toString() == 'SHAPE') {
 
      const contents = slides[i].getPageElements()[j].asShape().getText().asString();
     
      /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
      const notes    = slides[i].getNotesPage();
      notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
      /*翻訳後のテキストを貼り付け*/
      const results  = LanguageApp.translate(contents,'en','ja');
      slides[i].getPageElements()[j].asShape().getText().setText(results);
    }
    else if (slides[i].getPageElements()[j].getPageElementType().toString() == 'GROUP'){
      const elemsInGroup = slides[i].getPageElements()[j].asGroup().getChildren()
      for(const elem of elemsInGroup) {
        /*console.log(elem.getPageElementType().toString());*/
        if (elem.getPageElementType().toString() === 'SHAPE') {
          const contents = elem.asShape().getText().asString();
     
          /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
          const notes    = slides[i].getNotesPage();
          notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
          /*翻訳後のテキストを貼り付け*/
          const results  = LanguageApp.translate(contents,'en','ja');
          elem.asShape().getText().setText(results);
        }
        else if (elem.getPageElementType().toString() === 'GROUP') {
          const elemsInGroup2 = elem.asGroup().getChildren()
          for(const elem2 of elemsInGroup2) {
            console.log(elem2.getPageElementType().toString());
            if (elem2.getPageElementType().toString() === 'SHAPE') {
              const contents = elem2.asShape().getText().asString();
     
              /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
              const notes    = slides[i].getNotesPage();
              notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
     
              /*翻訳後のテキストを貼り付け*/
              const results  = LanguageApp.translate(contents,'en','ja');
              elem2.asShape().getText().setText(results);
            }
          }
        }
      }
    }
  }//for_j
}//for_i
/*テーブル内も変換出来るようにする*//*全スライドの取得*/
for(let i = 0; i < slides.length; i++){
  for(let j = 0; j < slides[i].getTables().length; j++){
     
    console.log('スライド　%s　ページ目には、表　%sがあります。',i+1,slides[i].getTables());
     
    const rowIndex    = slides[i].getTables()[j].getNumRows();
    const columnIndex = slides[i].getTables()[j].getNumColumns();
     
    console.log('行数：　%s',rowIndex);
    console.log('列数：　%s',columnIndex);
     
    for(let k = 0; k < rowIndex; k++){
      for(let l = 0; l < columnIndex; l++){
         
        /**/
        const innertable  = slides[i].getTables()[j].getCell(k, l).getText().asString();
        const results     = LanguageApp.translate(innertable,'en','ja');
         
        console.log('getCell(%s, %s)　元のテキスト：　%S',k,l,innertable);
        console.log('getCell(%s, %s)　翻訳後のテキスト：　%S',k,l,results);
         
         
        if(innertable === ''){continue}
        slides[i].getTables()[j].getCell(k, l).getText().setText(results);
      }//for_l
    }//for_k
  }//for_j
}//for_i
}//end
 
 
/*特定のページに存在する表内のテキストを翻訳する*/
function specifiedPageTranslate() {
const presentation  = SlidesApp.getActivePresentation();
const slides        = presentation.getSlides();
 
/*画面の操作を変化させるための記述*/
const ui            = SlidesApp.getUi();
const response      = ui.prompt(
  '翻訳こんにゃく',
  '翻訳したいスライドのページ番号を入力してください。',
  ui.ButtonSet.OK_CANCEL
);
 
//スライドページは配列で取得されるため
let page = response.getResponseText();
page -= 1;
 
switch(response.getSelectedButton()){
  case ui.Button.OK:
    console.log('ページ番号は、%s です。', page);
    break;
  
  case ui.Button.CANCEL:
    console.log('キャンセルが押されたため、処理を中断しました。');
    break;
     
  case ui.Button.CLOSE:
    console.log('閉じるボタンが押されました。');
}
 
for(let j = 0; j < slides[page].getPageElements().length; j++){
  /*console.log(slides[page].getPageElements()[j].getPageElementType().toString())*/
  if (slides[page].getPageElements()[j].getPageElementType().toString() == 'SHAPE') {
 
    const contents = slides[page].getPageElements()[j].asShape().getText().asString();
   
    /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
    const notes    = slides[page].getNotesPage();
    notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
    /*翻訳後のテキストを貼り付け*/
    const results  = LanguageApp.translate(contents,'en','ja');
    slides[page].getPageElements()[j].asShape().getText().setText(results);
  }
  else if (slides[page].getPageElements()[j].getPageElementType().toString() == 'GROUP'){
    const elemsInGroup = slides[page].getPageElements()[j].asGroup().getChildren()
    for(const elem of elemsInGroup) {
      /*console.log(elem.getPageElementType().toString());*/
      if (elem.getPageElementType().toString() === 'SHAPE') {
        const contents = elem.asShape().getText().asString();
   
        /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
        const notes    = slides[page].getNotesPage();
        notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
        /*翻訳後のテキストを貼り付け*/
        const results  = LanguageApp.translate(contents,'en','ja');
        elem.asShape().getText().setText(results);
      }
      else if (elem.getPageElementType().toString() === 'GROUP') {
        const elemsInGroup2 = elem.asGroup().getChildren()
        for(const elem2 of elemsInGroup2) {
          console.log(elem2.getPageElementType().toString());
          if (elem2.getPageElementType().toString() === 'SHAPE') {
            const contents = elem2.asShape().getText().asString();
   
            /*スピーカーノートの取得 翻訳前のテキストをスピーカノートに入れる*/
            const notes    = slides[page].getNotesPage();
            notes.getSpeakerNotesShape().getText().appendText("\n" + contents);
   
            /*翻訳後のテキストを貼り付け*/
            const results  = LanguageApp.translate(contents,'en','ja');
            elem2.asShape().getText().setText(results);
          }
        }
      }
    }
  }
}//for_j
 
/*特定のページ番号を取得*/
for(let i = 0; i < slides[page].getTables().length; i++){
   
  const rowIndex    = slides[page].getTables()[i].getNumRows();
  const columnIndex = slides[page].getTables()[i].getNumColumns();
   
  console.log('行数：　%s',rowIndex);
  console.log('列数：　%s',columnIndex);
   
  for(let j = 0; j < rowIndex; j++){
    for(let k = 0; k < columnIndex; k++){
       
      /**/
      const innertable  = slides[page].getTables()[i].getCell(j, k).getText().asString();
      const results     = LanguageApp.translate(innertable,'en','ja');
       
      console.log('getCell(%s, %s)　元のテキスト：　%S',j,k,innertable);
      console.log('getCell(%s, %s)　翻訳後のテキスト：　%S',j,k,results);
       
      if(innertable === ''){continue}
      slides[page].getTables()[i].getCell(j, k).getText().setText(results);
    }//for_k
  }//for_j
}//for_i
}//end
