---
title: 検索オプションを使用して Word アドインでテキストを検索する
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: ca5c819edb7f3c183379d9df997e41eb56a4de51
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505371"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>検索オプションを使用して Word アドインでテキストを検索する 

アドインはドキュメントのテキストに対する処理を行う必要が頻繁にあります。検索機能はすべてのコンテンツ コントロール ( [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js)、[Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js)、[Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js)、[Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js)、[TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)、ベースの [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) オブジェクトなど) で利用できます。この関数は、検索するテキストを表す文字列 (またはワイルドカード) と [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) オブジェクトを引数とし、検索したテキストと一致する範囲のコレクションを返します。

## <a name="search-options"></a>検索オプション
検索オプションは、検索パラメータの処理方法を定義するブール値のコレクションです。 

| プロパティ     | 説明|
|:---------------|:----|
|ignorePunct|単語間のすべての句読点を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに対応します。|
|ignoreSpace|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに対応します。|
|matchCase|大文字と小文字を区別して検索するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [大文字と小文字を区別する」チェック ボックスに対応します。|
|matchPrefix|検索する文字列で始まる単語と一致する検索を行うかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接頭辞に一致する] チェック ボックスに対応します。|
|matchSuffix|検索する文字列で終わる単語と一致する検索を行うかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接尾辞に一致する] チェックボックスに対応します。|
|matchWholeWord|長い単語の一部ではなく、全体が一致する単語のみを検索の対象にするかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [完全に一致する単語だけを検索する] チェック ボックスに対応します。|
|matchWildcards|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに対応します。|

## <a name="wildcard-guidance"></a>ワイルドカード ガイダンス
次の表は、Word JavaScript API の検索ワイルドカードに関するガイダンスを示しています。

| 検索対象：         | ワイルドカード |  サンプル |
|:-----------------|:--------|:----------|
| 任意の 1 文字| ? |s?t は、sat や set を検出します。 |
|1 文字以上の任意の文字列| * |s*d は、sad や started を検出します。|
|単語の先頭|< |<(inter) は、interesting や intercept を検出しますが、splintered は検出しません。|
|単語の末尾 |> |(in)> は、in や within を検出しますが、interesting は検出しません。|
|指定した文字のいずれか 1 つ|[ ] |w[io]n は、win と won を検出します。|
|指定した範囲内の任意の 1 文字| [-] |[r-t]ight は、right や sight を検出します。範囲は、昇順で指定します。|
|[ ] 内に指定した範囲以外の任意の 1 文字|[!x-z] |t[!a-m]ck は、tock や tuck を検出しますが、tack や tick は検出しません。|
|直前の文字または式の n 回の繰り返し|{n} |fe{2}d は、feed を検出しますが、fed は検出しません。|
|直前の文字または式の n 回以上の繰り返し|{n,} |fe{1,}d は、fed および feed を検出します。|
|直前の文字または式の n 回以上 m 回以下の繰り返し|{n,m} |10{1,3}は、10、100、1000 を検出します。|
|直前の文字または式の 1 回以上の繰り返し|@ |lo@t は、lot や loot を検出します。|

### <a name="escaping-the-special-characters"></a>特殊文字のエスケープ

ワイルドカード検索は、正規表現による検索と基本的には同じです。正規表現には、'['、']'、'('、')'、'{'、'}'、'\*'、'?'、'<'、'>'、'!'、'@' などの特殊文字があります。コードが検索する文字列リテラルの一部が特殊文字の場合は、特殊文字を正規表現のロジックの一部としてではなく、単なる文字として扱う必要があることを Word が認識できるようにエスケープする必要があります。Word UI 検索で文字をエスケープするには、その文字の前に '\' を付けます。ただしプログラムの中でエスケープする場合は、'[]' の間に配置します。たとえば、'[\*]\*' は、'\*' で始まり、その後に任意の数の他の文字が続く文字列を検索します。 

## <a name="examples"></a>例
次の例は、一般的なシナリオを示しています。

### <a name="ignore-punctuation-search"></a>次の例は、一般的なシナリオを示しています。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a>接頭辞に基づく検索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a>接尾辞に基づく検索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a>ワイルドカードを使用する検索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

詳しい情報は、[Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js) で検索できます。