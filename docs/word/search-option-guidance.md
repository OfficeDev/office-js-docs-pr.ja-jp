---
title: 検索オプションを使用して Word アドインでテキストを検索する
description: ''
ms.date: 07/20/2018
ms.openlocfilehash: d2c0fa2d542cd64986c2fd82f8a50a813f14610a
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270622"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="676dc-102">検索オプションを使用して Word アドインでテキストを検索する</span><span class="sxs-lookup"><span data-stu-id="676dc-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="676dc-103">アドインは、ドキュメントのテキストに基づいて動作することが必要な場合がよくあります。</span><span class="sxs-lookup"><span data-stu-id="676dc-103">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="676dc-104">検索関数は、各コンテンツ コントロール (これには、[Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js)、[Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js)、[Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js)、[Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js)、[TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)、およびベース [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) オブジェクトが含まれます) で公開されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-104">A search function is exposed by every content control (this includes [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js), and the base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) object).</span></span> <span data-ttu-id="676dc-105">この関数には、検索しているテキストおよび [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) オブジェクトを表す文字列 (またはワイルドカード式) を使用します。</span><span class="sxs-lookup"><span data-stu-id="676dc-105">This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) object.</span></span> <span data-ttu-id="676dc-106">これにより、検索テキストと一致する範囲のコレクションが返されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-106">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="676dc-107">検索オプション</span><span class="sxs-lookup"><span data-stu-id="676dc-107">Search options</span></span>
<span data-ttu-id="676dc-108">検索オプションとは、検索パラメーターをどのように処理するかを定義するブール値のコレクションです。</span><span class="sxs-lookup"><span data-stu-id="676dc-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="676dc-109">プロパティ</span><span class="sxs-lookup"><span data-stu-id="676dc-109">Property</span></span>     | <span data-ttu-id="676dc-110">説明</span><span class="sxs-lookup"><span data-stu-id="676dc-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="676dc-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="676dc-111">ignorePunct</span></span>|<span data-ttu-id="676dc-112">単語間の句読点文字をすべて無視するかどうかを示す値を取得するか設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-112">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="676dc-113">[検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに相当します。</span><span class="sxs-lookup"><span data-stu-id="676dc-113">True ignores all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="676dc-114">ignoreSpace</span></span>|<span data-ttu-id="676dc-115">単語間のすべての空白を無視するかどうかを示す値を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-115">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="676dc-116">[検索と置換] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに相当します。</span><span class="sxs-lookup"><span data-stu-id="676dc-116">Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="676dc-117">matchCase</span></span>|<span data-ttu-id="676dc-118">大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-118">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="676dc-119">[検索と置換] ダイアログ ボックスの [大文字と小文字を区別する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="676dc-119">Corresponds to the Match case check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="676dc-120">matchPrefix</span></span>|<span data-ttu-id="676dc-121">検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-121">Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="676dc-122">[検索と置換] ダイアログ ボックスの [接頭辞に一致する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="676dc-122">Corresponds to the Match prefix check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="676dc-123">matchSuffix</span></span>|<span data-ttu-id="676dc-124">検索文字列で終わる単語と一致するかどうかを示す値を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-124">Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="676dc-125">[検索と置換] ダイアログ ボックスの [接尾辞に一致する] に対応します。</span><span class="sxs-lookup"><span data-stu-id="676dc-125">Corresponds to the Match suffix check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="676dc-126">matchWholeWord</span></span>|<span data-ttu-id="676dc-127">長い単語の一部のテキストではなく、単語全体のみを検索するかどうかを示す値を取得するか設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-127">Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="676dc-128">[検索と置換] ダイアログ ボックスの [完全に一致する単語だけを検索する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="676dc-128">Corresponds to the Find whole words only check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="676dc-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="676dc-129">matchWildcards</span></span>|<span data-ttu-id="676dc-130">特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="676dc-130">Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="676dc-131">[検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに対応します。</span><span class="sxs-lookup"><span data-stu-id="676dc-131">Corresponds to the Use wildcards check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="676dc-132">ワイルドカードに関する説明</span><span class="sxs-lookup"><span data-stu-id="676dc-132">Wildcard Guidance</span></span>
<span data-ttu-id="676dc-133">次の表では、Word JavaScript API の検索ワイルドカードについて説明します。</span><span class="sxs-lookup"><span data-stu-id="676dc-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="676dc-134">検索方法</span><span class="sxs-lookup"><span data-stu-id="676dc-134">To find:</span></span>         | <span data-ttu-id="676dc-135">ワイルドカード</span><span class="sxs-lookup"><span data-stu-id="676dc-135">Wildcard</span></span> |  <span data-ttu-id="676dc-136">サンプル</span><span class="sxs-lookup"><span data-stu-id="676dc-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="676dc-137">任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="676dc-137">Any single character</span></span>| <span data-ttu-id="676dc-138">?</span><span class="sxs-lookup"><span data-stu-id="676dc-138"></span></span> |<span data-ttu-id="676dc-139">s?t は、sat や set を検出します。</span><span class="sxs-lookup"><span data-stu-id="676dc-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="676dc-140">文字からなる任意の文字列</span><span class="sxs-lookup"><span data-stu-id="676dc-140">Any string of characters</span></span>| * |<span data-ttu-id="676dc-141">s\*d は、sad や started を検出します。</span><span class="sxs-lookup"><span data-stu-id="676dc-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="676dc-142">単語の先頭</span><span class="sxs-lookup"><span data-stu-id="676dc-142">The beginning of a word</span></span>|< |<span data-ttu-id="676dc-143"><(inter) では、interesting や intercept が検出されますが、splintered は検出されません。</span><span class="sxs-lookup"><span data-stu-id="676dc-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="676dc-144">単語の末尾</span><span class="sxs-lookup"><span data-stu-id="676dc-144">The end of a word</span></span> |> |<span data-ttu-id="676dc-145">(in)> では、in や within が検出されますが、interesting は検出されません。</span><span class="sxs-lookup"><span data-stu-id="676dc-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="676dc-146">指定した文字のいずれか 1 つ</span><span class="sxs-lookup"><span data-stu-id="676dc-146">One of the specified characters</span></span>|<span data-ttu-id="676dc-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="676dc-147"></span></span> |<span data-ttu-id="676dc-148">w[io]n では、win と won が検出されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="676dc-149">この範囲に含まれる任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="676dc-149">Any single character in this range</span></span>| <span data-ttu-id="676dc-150">[-]</span><span class="sxs-lookup"><span data-stu-id="676dc-150"></span></span> |<span data-ttu-id="676dc-p109">[r-t]ight では、right や sight が検出されます。範囲は、昇順にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="676dc-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="676dc-153">角括弧で囲まれた範囲の文字を除く任意の 1 文字</span><span class="sxs-lookup"><span data-stu-id="676dc-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="676dc-155">t[!a-m]ck では、tock や tuck が検出されますが、tack や tick は検出されません。</span><span class="sxs-lookup"><span data-stu-id="676dc-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="676dc-156">直前の文字または式の n 回の出現</span><span class="sxs-lookup"><span data-stu-id="676dc-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="676dc-157">{n}</span><span class="sxs-lookup"><span data-stu-id="676dc-157">{n}</span></span> |<span data-ttu-id="676dc-158">fe{2}d では、feed が検出されますが、fed は検出されません。</span><span class="sxs-lookup"><span data-stu-id="676dc-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="676dc-159">直前の文字または式の n 回以上の出現</span><span class="sxs-lookup"><span data-stu-id="676dc-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="676dc-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="676dc-160">{n,}</span></span> |<span data-ttu-id="676dc-161">fe{1,}d では、fed や feed が検出されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="676dc-162">直前の文字または式の n 回から m 回までの出現</span><span class="sxs-lookup"><span data-stu-id="676dc-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="676dc-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="676dc-163">{n,m}</span></span> |<span data-ttu-id="676dc-164">10{1,3} では、10、100、1000 が検出されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="676dc-165">直前の文字または式の 1 回以上の出現</span><span class="sxs-lookup"><span data-stu-id="676dc-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="676dc-166">lo@t では、lot や loot が検出されます。</span><span class="sxs-lookup"><span data-stu-id="676dc-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="676dc-167">特殊文字のエスケープ</span><span class="sxs-lookup"><span data-stu-id="676dc-167">Escaping the special characters</span></span>

<span data-ttu-id="676dc-p110">ワイルドカード検索は、基本的に正規表現での検索と同じです。正規表現には、'['、']'、'('、')'、'{'、'}'、'\*'、'?'、'<'、'>'、'!'、および '@' を含む特殊文字があります。これらの文字のいずれかが、コードが検索しているリテラル文字列の一部である場合は、その文字を正規表現のロジックの一部としてではなく、文字どおりに扱う必要があることを Word が認識できるように、エスケープする必要があります。Word UI 検索で文字をエスケープするには、その文字の前に '\' を付けます。ただしプログラムを使用してエスケープするには、これを '[]' 文字の間に配置します。たとえば、'[\*]\*' は、'\*' で始まり、その後に任意の数の他の文字が続く文字列を検索します。</span><span class="sxs-lookup"><span data-stu-id="676dc-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="676dc-173">例</span><span class="sxs-lookup"><span data-stu-id="676dc-173">Examples</span></span>
<span data-ttu-id="676dc-174">次の例では、よくあるシナリオについて説明します。</span><span class="sxs-lookup"><span data-stu-id="676dc-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="676dc-175">句読点を無視する検索</span><span class="sxs-lookup"><span data-stu-id="676dc-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="676dc-176">接頭辞に基づく検索</span><span class="sxs-lookup"><span data-stu-id="676dc-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="676dc-177">接尾辞に基づく検索</span><span class="sxs-lookup"><span data-stu-id="676dc-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="676dc-178">ワイルドカードを使用する検索</span><span class="sxs-lookup"><span data-stu-id="676dc-178">Search using a wildcard</span></span>

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

<span data-ttu-id="676dc-179">詳細については、「[Word JavaScript API の概要](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="676dc-179">More information can be found in the [Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span></span>