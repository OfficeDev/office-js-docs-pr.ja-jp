---
title: 検索オプションを使用して Word アドインでテキストを検索する
description: Word アドインで検索オプションを使用する方法について説明します。
ms.date: 02/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: e8f9dd2605af9307a49fabfafdecb0df4e97fe9f
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187344"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>検索オプションを使用して Word アドインでテキストを検索する

アドインは、ドキュメントのテキストに基づいて動作することが必要な場合がよくあります。
検索関数は、各コンテンツ コントロール (これには、[Body](/javascript/api/word/word.body)、[Paragraph](/javascript/api/word/word.paragraph)、[Range](/javascript/api/word/word.range)、[Table](/javascript/api/word/word.table)、[TableRow](/javascript/api/word/word.tablerow)、およびベース [ContentControl](/javascript/api/word/word.contentcontrol) オブジェクトが含まれます) で公開されます。 この関数には、検索しているテキストおよび [SearchOptions](/javascript/api/word/word.searchoptions) オブジェクトを表す文字列 (またはワイルドカード式) を使用します。 これにより、検索テキストと一致する範囲のコレクションが返されます。

## <a name="search-options"></a>検索オプション

検索オプションとは、検索パラメーターをどのように処理するかを定義するブール値のコレクションです。

| プロパティ       | 説明|
|:---------------|:----|
|ignorePunct|単語間の句読点文字をすべて無視するかどうかを示す値を取得するか設定します。 **[検索と置換**] ダイアログ ボックスの [句読点文字を無視する] チェック ボックスに対応します。|
|ignoreSpace|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。 **[検索と置換**] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに対応します。|
|matchCase|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。 **[検索と置換**] ダイアログ ボックスの [マッチ ケース] チェック ボックスに対応します。|
|matchPrefix|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。 **[検索と置換**] ダイアログ ボックスの [一致するプレフィックス] チェック ボックスに対応します。|
|matchSuffix|検索文字列で終わる単語と一致するかどうかを示す値を取得または設定します。 **[検索と置換**] ダイアログ ボックスの [一致するサフィックス] チェック ボックスに対応します。|
|matchWholeWord|長い単語の一部のテキストではなく、単語全体のみを検索するかどうかを示す値を取得するか設定します。 [検索と置換] ダイアログ ボックスの [単語全体のみ検索] チェック ボックス **に** 対応します。|
|matchWildcards|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。 **[検索と置換**] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに対応します。|

## <a name="wildcard-guidance"></a>ワイルドカードに関する説明

次の表では、Word JavaScript API の検索ワイルドカードに関するガイダンスを示します。

| 検索方法         | ワイルドカード |  サンプル |
|:-----------------|:--------|:----------|
|任意の 1 文字| ? |s?t は、sat や set を検出します。 |
|文字からなる任意の文字列| * |s*d は、sad や started を検出します。|
|単語の先頭|< |<(inter) では、interesting や intercept が検出されますが、splintered は検出されません。|
|単語の末尾 |> |(in)> では、in や within が検出されますが、interesting は検出されません。|
|指定した文字のいずれか 1 つ|[ ] |w[io]n では、win と won が検出されます。|
|この範囲に含まれる任意の 1 文字| [-] |[r-t]ight では、right や sight が検出されます。範囲は、昇順にする必要があります。|
|角括弧で囲まれた範囲の文字を除く任意の 1 文字|[!x-z] |t[!a-m]ck では、tock や tuck が検出されますが、tack や tick は検出されません。|
|直前の文字または式の正確に *n 個* の出現|{n} |fe{2}d では、feed が検出されますが、fed は検出されません。|
|前の文字または式の少なくとも *n 個* の出現回数|{n,} |fe{1,}d では、fed や feed が検出されます。|
|前の文字または式の *n* から *m* 個の出現回数|{n,m} |10{1,3} では、10、100、1000 が検出されます。|
|直前の文字または式の 1 回以上の出現|@ |lo@t では、lot や loot が検出されます。|

### <a name="escaping-special-characters"></a>特殊文字のエスケープ

ワイルドカード検索は、基本的に正規表現での検索と同じです。 正規表現には、'[''、']、'(''、')'、'{'、'}'、'、'\*?'、'<'、'>'、'!'、'@' などの特殊文字があります。 これらの文字の 1 つがコードが検索しているリテラル文字列の一部である場合は、正規表現のロジックの一部としてではなく、リテラルで扱う必要があることを Word が認識できるように、エスケープする必要があります。 Word UI 検索で文字をエスケープするには、前に円記号 (''\\) を付けますが、プログラムでエスケープするには、'[]' 文字の間に配置します。 たとえば、'[\*]' は\*、'' で\*始まる任意の文字列の後に任意の数の他の文字を検索します。

## <a name="examples"></a>例

次の例では、よくあるシナリオについて説明します。

### <a name="ignore-punctuation-search"></a>句読点を無視する検索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-prefix"></a>接頭辞に基づく検索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-suffix"></a>接尾辞に基づく検索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-using-a-wildcard"></a>ワイルドカードを使用する検索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

詳細については、「[Word JavaScript API の概要](../reference/overview/word-add-ins-reference-overview.md)」を参照してください。
