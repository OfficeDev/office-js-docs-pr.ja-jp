---
title: 検索オプションを使用して Word アドインでテキストを検索する
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: d81ffdcec49d59c175c3e5ecdf82ad1f796fdb3e
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944101"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>検索オプションを使用して Word アドインでテキストを検索する 

アドインは頻繁に文書のテキストに基づいて動作する必要があります。
検索機能はすべてのコンテンツ コントロールによって公開されます（これには [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js)、[Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js)、[Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js)、[Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js)、[TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js) およびベース [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) オブジェクトが含まれます）。 この関数は、検索しているテキストを表す文字列（またはwldcard 式）と [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) オブジェクトをとります。 検索テキストと一致する範囲のコレクションを返します。

## <a name="search-options"></a>検索オプション
検索オプションは、検索パラメータの処理方法を定義するブール値のコレクションです。 

| プロパティ     | 説明|
|:---------------|:----|
|ignorePunct|単語間のすべての句読点を無視するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに対応します。|
|ignoreSpace|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに対応します。|
|matchCase|大文字と小文字を区別して検索するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [大文字と小文字を区別] チェック ボックスに対応します。|
|matchPrefix|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [プレフィックスに一致] チェック ボックスに対応します。|
|matchSuffix|検索文字列で終わる単語と一致するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [サフィックスに一致] チェック ボックスに対応します。|
|matchWholeWord|長い単語の一部をなすテキストではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [単語全体のみを検索する] チェック ボックスに対応します。|
|matchWildcards|特別な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに対応します。|

## <a name="wildcard-guidance"></a>ワイルドカードのガイダンス
次の表は、Word JavaScript API の検索ワイルドカードに関するガイダンスを示しています。

| 検索対象         | ワイルドカード |  サンプル |
|:-----------------|:--------|:----------|
| 任意の 1 文字| ? |s?t は、sat や set を検出します。 |
|文字からなる任意の文字列| * |s*d は、sad や started を検出します。|
|単語の先頭|< |<(inter) では、interesting や intercept が検出されますが、splintered は検出されません。|
|単語の末尾 |> |(in)> では、in や within が検出されますが、interesting は検出されません。|
|指定した文字のいずれか 1 つ|[ ] |w[io]n では、win と won が検出されます。|
|この範囲に含まれる任意の 1 文字| [-] |[r-t]ight では、right や sight が検出されます。範囲は、昇順にする必要があります。|
|角括弧で囲まれた範囲の文字を除く任意の 1 文字|[!x-z] |t[!a-m]ck では、tock や tuck が検出されますが、tack や tick は検出されません。|
|直前の文字または式の n 回の出現|{n} |fe{2}d では、feed が検出されますが、fed は検出されません。|
|直前の文字または式の n 回以上の出現|{n,} |fe{1,}d では、fed および feed が検出されます。|
|直前の文字または式の n 回から m 回までの出現|{n,m} |10{1,3}では 10、100、1000 が検出されます。|
|直前の文字または式の 1 回以上の出現|@ |lo@t では、lot や loot が検出されます。|

### <a name="escaping-the-special-characters"></a>特殊文字のエスケープ

ワイルドカード検索は、基本的に正規表現での検索と同じです。正規表現には、'['、']'、'('、')'、'{'、'}'、'\*'、'?'、'<'、'>'、'!'、および '@' を含む特殊文字があります。これらの文字のいずれかが、コードが検索しているリテラル文字列の一部である場合は、その文字を正規表現のロジックの一部としてではなく、文字どおりに扱う必要があることを Word が認識できるように、エスケープする必要があります。Word UI 検索で文字をエスケープするには、その文字の前に '\' を付けます。ただしプログラムを使用してエスケープするには、これを '[]' 文字の間に配置します。たとえば、'[\*]\*' は、'\*' で始まり、その後に任意の数の他の文字が続く文字列を検索します。 

## <a name="examples"></a>例
次の例は、一般的なシナリオを示しています。

### <a name="ignore-punctuation-search"></a>句読点を無視する検索

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

詳しい情報は、[Word JavaScript Reference API](https://docs.microsoft.com/javascript/office/overview/word-add-ins-reference-overview?view=office-js) で検索できます。