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
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="07854-102">検索オプションを使用して Word アドインでテキストを検索する</span><span class="sxs-lookup"><span data-stu-id="07854-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="07854-p101">アドインはドキュメントのテキストに対する処理を行う必要が頻繁にあります。検索機能はすべてのコンテンツ コントロール ( [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js)、[Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js)、[Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js)、[Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js)、[TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)、ベースの [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) オブジェクトなど) で利用できます。この関数は、検索するテキストを表す文字列 (またはワイルドカード) と [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) オブジェクトを引数とし、検索したテキストと一致する範囲のコレクションを返します。</span><span class="sxs-lookup"><span data-stu-id="07854-p101">Add-ins frequently need to act based on the text of a document. A search function is exposed by every content control (this includes [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js), and the base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) object). This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) object. It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="07854-107">検索オプション</span><span class="sxs-lookup"><span data-stu-id="07854-107">Search options</span></span>
<span data-ttu-id="07854-108">検索オプションは、検索パラメータの処理方法を定義するブール値のコレクションです。</span><span class="sxs-lookup"><span data-stu-id="07854-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="07854-109">プロパティ</span><span class="sxs-lookup"><span data-stu-id="07854-109">Property</span></span>     | <span data-ttu-id="07854-110">説明</span><span class="sxs-lookup"><span data-stu-id="07854-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="07854-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="07854-111">ignorePunct</span></span>|<span data-ttu-id="07854-p102">単語間のすべての句読点を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p102">Gets or sets a value indicating whether to ignore all punctuation characters between words. Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="07854-114">ignoreSpace</span></span>|<span data-ttu-id="07854-p103">単語間のすべての空白を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p103">Gets or sets a value indicating whether to ignore all whitespace between words. Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="07854-117">matchCase</span></span>|<span data-ttu-id="07854-p104">大文字と小文字を区別して検索するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [大文字と小文字を区別する」チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p104">Gets or sets a value indicating whether to perform a case sensitive search. Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="07854-120">matchPrefix</span></span>|<span data-ttu-id="07854-p105">検索する文字列で始まる単語と一致する検索を行うかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接頭辞に一致する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p105">Gets or sets a value indicating whether to match words that begin with the search string. Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="07854-123">matchSuffix</span></span>|<span data-ttu-id="07854-p106">検索する文字列で終わる単語と一致する検索を行うかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接尾辞に一致する] チェックボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p106">Gets or sets a value indicating whether to match words that end with the search string. Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="07854-126">matchWholeWord</span></span>|<span data-ttu-id="07854-p107">長い単語の一部ではなく、全体が一致する単語のみを検索の対象にするかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [完全に一致する単語だけを検索する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p107">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word. Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="07854-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="07854-129">matchWildcards</span></span>|<span data-ttu-id="07854-p108">特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="07854-p108">Gets or sets a value indicating whether the search will be performed using special search operators. Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="07854-132">ワイルドカード ガイダンス</span><span class="sxs-lookup"><span data-stu-id="07854-132">Wildcard Guidance</span></span>
<span data-ttu-id="07854-133">次の表は、Word JavaScript API の検索ワイルドカードに関するガイダンスを示しています。</span><span class="sxs-lookup"><span data-stu-id="07854-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="07854-134">検索対象：</span><span class="sxs-lookup"><span data-stu-id="07854-134">To find:</span></span>         | <span data-ttu-id="07854-135">ワイルドカード</span><span class="sxs-lookup"><span data-stu-id="07854-135">Wildcard</span></span> |  <span data-ttu-id="07854-136">サンプル</span><span class="sxs-lookup"><span data-stu-id="07854-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="07854-137">任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="07854-137">Any single character</span></span>| <span data-ttu-id="07854-138">?</span><span class="sxs-lookup"><span data-stu-id="07854-138">?</span></span> |<span data-ttu-id="07854-139">s?t は、sat や set を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="07854-140">1 文字以上の任意の文字列</span><span class="sxs-lookup"><span data-stu-id="07854-140">Any string of characters</span></span>| * |<span data-ttu-id="07854-141">s\*d は、sad や started を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="07854-142">単語の先頭</span><span class="sxs-lookup"><span data-stu-id="07854-142">The beginning of a word</span></span>|< |<span data-ttu-id="07854-143"><(inter) は、interesting や intercept を検出しますが、splintered は検出しません。</span><span class="sxs-lookup"><span data-stu-id="07854-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="07854-144">単語の末尾</span><span class="sxs-lookup"><span data-stu-id="07854-144">The end of a word</span></span> |> |<span data-ttu-id="07854-145">(in)> は、in や within を検出しますが、interesting は検出しません。</span><span class="sxs-lookup"><span data-stu-id="07854-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="07854-146">指定した文字のいずれか 1 つ</span><span class="sxs-lookup"><span data-stu-id="07854-146">One of the specified characters</span></span>|<span data-ttu-id="07854-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="07854-147">[ ]</span></span> |<span data-ttu-id="07854-148">w[io]n は、win と won を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="07854-149">指定した範囲内の任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="07854-149">Any single character in this range</span></span>| <span data-ttu-id="07854-150">[-]</span><span class="sxs-lookup"><span data-stu-id="07854-150">[-]</span></span> |<span data-ttu-id="07854-p109">[r-t]ight は、right や sight を検出します。範囲は、昇順で指定します。</span><span class="sxs-lookup"><span data-stu-id="07854-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="07854-153">[ ] 内に指定した範囲以外の任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="07854-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="07854-155">t[!a-m]ck は、tock や tuck を検出しますが、tack や tick は検出しません。</span><span class="sxs-lookup"><span data-stu-id="07854-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="07854-156">直前の文字または式の n 回の繰り返し</span><span class="sxs-lookup"><span data-stu-id="07854-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="07854-157">{n}</span><span class="sxs-lookup"><span data-stu-id="07854-157">{n}</span></span> |<span data-ttu-id="07854-158">fe{2}d は、feed を検出しますが、fed は検出しません。</span><span class="sxs-lookup"><span data-stu-id="07854-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="07854-159">直前の文字または式の n 回以上の繰り返し</span><span class="sxs-lookup"><span data-stu-id="07854-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="07854-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="07854-160">{n,}</span></span> |<span data-ttu-id="07854-161">fe{1,}d は、fed および feed を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="07854-162">直前の文字または式の n 回以上 m 回以下の繰り返し</span><span class="sxs-lookup"><span data-stu-id="07854-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="07854-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="07854-163">{n,m}</span></span> |<span data-ttu-id="07854-164">10{1,3}は、10、100、1000 を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="07854-165">直前の文字または式の 1 回以上の繰り返し</span><span class="sxs-lookup"><span data-stu-id="07854-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="07854-166">lo@t は、lot や loot を検出します。</span><span class="sxs-lookup"><span data-stu-id="07854-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="07854-167">特殊文字のエスケープ</span><span class="sxs-lookup"><span data-stu-id="07854-167">Escaping the special characters</span></span>

<span data-ttu-id="07854-p110">ワイルドカード検索は、正規表現による検索と基本的には同じです。正規表現には、'['、']'、'('、')'、'{'、'}'、'\*'、'?'、'<'、'>'、'!'、'@' などの特殊文字があります。コードが検索する文字列リテラルの一部が特殊文字の場合は、特殊文字を正規表現のロジックの一部としてではなく、単なる文字として扱う必要があることを Word が認識できるようにエスケープする必要があります。Word UI 検索で文字をエスケープするには、その文字の前に '\' を付けます。ただしプログラムの中でエスケープする場合は、'[]' の間に配置します。たとえば、'[\*]\*' は、'\*' で始まり、その後に任意の数の他の文字が続く文字列を検索します。</span><span class="sxs-lookup"><span data-stu-id="07854-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="07854-173">例</span><span class="sxs-lookup"><span data-stu-id="07854-173">Examples</span></span>
<span data-ttu-id="07854-174">次の例は、一般的なシナリオを示しています。</span><span class="sxs-lookup"><span data-stu-id="07854-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="07854-175">次の例は、一般的なシナリオを示しています。</span><span class="sxs-lookup"><span data-stu-id="07854-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="07854-176">接頭辞に基づく検索</span><span class="sxs-lookup"><span data-stu-id="07854-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="07854-177">接尾辞に基づく検索</span><span class="sxs-lookup"><span data-stu-id="07854-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="07854-178">ワイルドカードを使用する検索</span><span class="sxs-lookup"><span data-stu-id="07854-178">Search using a wildcard</span></span>

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

<span data-ttu-id="07854-179">詳しい情報は、[Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js) で検索できます。</span><span class="sxs-lookup"><span data-stu-id="07854-179">More information can be found in the [Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span></span>