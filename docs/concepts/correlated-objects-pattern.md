---
title: ループでのコンテキスト同期方法の使用を避ける
description: 分割ループと相関オブジェクトのパターンを使用して、コンテキストの呼び出しを回避する方法について説明します。ループでの同期。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 0f967b07b3ccf323321779676021c53c81102f83
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43226013"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>ループでのコンテキスト同期方法の使用を避ける

> [!NOTE]
> この記事では、バッチシステムを使用して Office ドキュメントを操作する、Excel、Word、OneNote、Visio&mdash;&mdash;用の4つのホスト固有の Office JavaScript api のうち、少なくとも1つを操作する最初の段階にとどまらないことを前提としています。 特に、呼び出し`context.sync`とは何か、コレクションオブジェクトについて理解しておく必要があることを理解しておく必要があります。 その段階になっていない場合は、「 [Office JAVASCRIPT API](../develop/understanding-the-javascript-api-for-office.md) 」と「ホスト固有」の下にリンクされているドキュメントを理解してください。

ホスト固有の API モデル (Excel、Word、OneNote、Visio) の1つを使用する Office アドインのプログラミングシナリオによっては、コードでコレクションオブジェクトのすべてのメンバーからいくつかのプロパティを読み取り、書き込み、または処理する必要があります。 たとえば、特定のテーブル列のすべてのセルの値を取得する必要がある Excel アドイン、またはドキュメント内の文字列のすべてのインスタンスを強調表示する必要がある Word アドイン。 コレクションオブジェクトの`items`プロパティのメンバーを反復処理する必要があります。ただし、パフォーマンス上の理由から、ループのすべて`context.sync`の反復処理で通話を回避する必要があります。 のすべての`context.sync`呼び出しでは、アドインから Office ドキュメントへのラウンドトリップがあります。 ラウンドトリップがインターネット経由で行われるため、特に web 上の Office でアドインが実行されている場合、ラウンドトリップが繰り返されるとパフォーマンスが低下します。

> [!NOTE]
> この記事のすべての例`for`ではループを使用していますが、ここで説明する操作は、次のような配列を反復処理できる任意のループステートメントに適用されます。
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> これらは、関数が渡され、配列内の項目に適用される、次のような配列メソッドにも適用されます。
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

## <a name="writing-to-the-document"></a>ドキュメントへの書き込み

最も単純なケースでは、コレクションオブジェクトのメンバーだけが書き込み、プロパティの読み取りは行われません。 たとえば、次のコードでは、Word 文書内の "the" のすべてのインスタンスが黄色で強調表示されています。 

> [!NOTE]
> 通常`context.sync`は、ホスト`run`メソッドの閉じる側の "}" 文字 ( `Excel.run` `Word.run`、など) の直前に配置することをお勧めします。 これは、メソッド`run`では、まだ同期さ`context.sync`れていないキューに入れられたコマンドがある場合に限り、最後に実行したときと同じ方法で非表示の呼び出しを行うためです。 この呼び出しが非表示になっていることがわかりやすいため、通常は明示的`context.sync`なを追加することをお勧めします。 ただし、この記事では、呼び出しを最小限に`context.sync`することについて説明していますが、 `context.sync`実際には不要な最終処理を追加する方が混乱しています。 そのため、この記事では、の最後に同期されていないコマンドがない場合`run`は、このままにします。 

```javascript
Word.run(async function (context) {
    let startTime, endTime;
    const docBody = context.document.body;

    // search() returns an array of Ranges.
    const searchResults = docBody.search('the', { matchWholeWord: true });
    context.load(searchResults, 'items');
    await context.sync();

    // Record the system time.
    startTime = performance.now();

    for (var i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].font.highlightColor = '#FFFF00';

      await context.sync(); // SYNCHRONIZE IN EACH ITERATION
    }
    
    // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

    // Record the system time again then calculate how long the operation took.
    endTime = performance.now();
    console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
  })
}
```

前のコードでは、Word の "the" という単語に200のインスタンスが含まれているドキュメントで完了するために1秒で完了していました。 しかし、ループ`await context.sync();`の内側の行がコメントアウトされているときに、ループがコメントアウトされた直後の行であれば、処理には1秒あたり10秒だけかかります。 Web 上の Word (ブラウザーとしてのエッジを含む) では、ループ内で同期が行われ、ループの後に同期が5倍高速になるまでに3秒で完了しました。 "The" の2000インスタンスが含まれるドキュメントでは、(web 上の Word) 80 秒で、ループの内部で同期が行われ、ループ後の同期では約20倍高速になりました。

> [!NOTE]
> 同期が同時に実行された場合に、ループ内での同期のバージョンが高速に実行されるかどうかを確認する必要`await`があります。これは`context.sync()`、の前にあるキーワードを削除するだけで行うことができます。 これにより、ランタイムが同期を開始し、同期が完了するのを待たずに、ループの次の反復処理を直ちに開始することができます。 ただし、次の理由から、ループを完全に移動`context.sync`するのではなく、この方法を使用しても問題ありません。
>
> - 同期バッチジョブのコマンドがキューに入れられているのと同様に、バッチジョブ自体は Office でキューに入れられますが、Office はキュー内に50個を超えるバッチジョブをサポートしません。 これ以上トリガーエラーが発生します。 そのため、ループに50を超える反復がある場合は、キューのサイズを超えている可能性があります。 反復回数が多いほど、発生する可能性が高くなります。 
> - "同時" は同時に意味がありません。 その場合でも、1つの同期操作を実行するよりも、複数の同期操作を実行するのにかかる時間が長くなります。
> - 同時操作は、開始したときと同じ順序で完了することは保証されません。 前の例では、単語 "the" が強調表示されている順序は重要ではありませんが、コレクション内のアイテムを順番に処理することが重要なシナリオがあります。

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a>分割ループパターンを使用してドキュメントから値を読み取る

ループ`context.sync`内のを回避することは、コードがそれぞれを処理するときにコレクションアイテムのプロパティを*読み取る*必要がある場合に、より困難になります。 コードで、Word 文書内のすべてのコンテンツコントロールを反復処理し、各コントロールに関連付けられている最初の段落のテキストをログに記録する必要があるとします。 プログラミングの instincts によって、コントロールをループ処理し、各`text` (最初の) 段落のプロパティを読み込ん`context.sync`で、プロキシの paragraph オブジェクトにドキュメントのテキストを設定する呼び出しを行い、それを記録することができます。 次に例を示します。

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

このシナリオでは、ループにが`context.sync`含まれないようにするために、**スプリットループ**パターンを呼び出すパターンを使用する必要があります。 このパターンの具体的な例を参照してから、正式な説明にしてみましょう。 スプリットループパターンを前述のコードスニペットに適用する方法は次のとおりです。 このコードについては、次の点に注意してください。

- これで2つのループが`context.sync`あり、どちらのループにも入っ`context.sync`ているので、一方のループ内にはありません。
- 最初のループでは、コレクションオブジェクトのアイテムを反復処理し`text` 、元のループと同じようにプロパティを読み込みますが、最初のループには、 `context.sync` `text` `paragraph`プロキシオブジェクトのプロパティを設定するためのが含まれていないため、段落テキストをログに記録することはできません。 代わりに、 `paragraph`オブジェクトを配列に追加します。
- 2番目のループでは、最初のループによって作成された`text`配列を`paragraph`反復処理し、各アイテムのをログに記録します。 これが可能なの`context.sync`は、2つのループの間に`text`あるがすべてのプロパティを設定したためです。

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

前の例では、を`context.sync`含むループをスプリットループパターンに入れるために、次の手順を提案しています。 

1. ループを2つのループに置き換えます。
2. 最初のループを作成してコレクションを反復処理し、各項目を配列に追加します。また、コードで読み込む必要のあるアイテムのプロパティも読み込みます。 
3. 最初のループに従って`context.sync` 、プロキシオブジェクトに読み込み済みのプロパティを設定するように呼び出します。 
4. 第 2 `context.sync`のループを使用して、最初のループで作成された配列に対して反復処理を行い、読み込まれたプロパティを読み取ります。

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a>[相関オブジェクト] パターンを使用してドキュメント内のオブジェクトを処理する

コレクション内のアイテムを処理するには、アイテム自体にはないデータが必要になるという、より複雑なシナリオを考えてみましょう。 このシナリオでは、テンプレートから作成されたドキュメントに対して何らかの定型テキストを使用して操作する Word アドインを小売します。 テキストに分散されているのは、"{コーディネーター}"、"{Deputy}"、および "{Manager}" の各プレースホルダー文字列の1つ以上のインスタンスです。 各プレースホルダーは、アドインによってユーザーの名前に置き換えられます。 この記事では、アドインの UI は重要ではありません。 たとえば、3つのテキストボックスを含む作業ウィンドウがあり、それぞれにプレースホルダーのいずれかでラベルが付けられているとします。 ユーザーは、各テキストボックスに名前を入力し、[**置換**] ボタンを押します。 ボタンのハンドラーは、名前をプレースホルダーにマップする配列を作成し、各プレースホルダーを割り当てられた名前に置き換えます。 

コードを試すために、この UI を使用してアドインを実際に作成する必要はありません。 [スクリプトラボツール](../overview/explore-with-script-lab.md)を使用して、重要なコードのプロトタイプを作成できます。 次の代入ステートメントを使用して、マッピング配列を作成します。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

次のコードは、内部ループを使用`context.sync`した場合に、各プレースホルダーを割り当てられた名前に置き換える方法を示しています。

```javascript
Word.run(async (context) => {

    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

上記のコードでは、外部ループと内部ループがあります。 各には、 `context.sync`が含まれています。 この記事の最初のコードスニペットに基づいて、内側のループの後`context.sync`に内側のループを単純に移動できることがわかるでしょう。 しかし、その場合でも、外側の`context.sync`ループには (実際にはそのうちの2つの) コードを残しておきます。 次のコードは、ループから削除`context.sync`する方法を示しています。 次のコードについて説明します。

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {        
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

メモコードでは、分割ループパターンを使用しています。

- 前の例の外側のループは、2つに分割されています。 (2 番目のループには内側のループがあります。これは、コードが一連のジョブ (またはプレースホルダー) を反復処理しており、そのセット内で一致する範囲を反復処理しているためです)。
- 各メジャーループ`context.sync`の後にはがあります`context.sync`が、ループ内にはありません。 
- 2番目のメジャーループでは、最初のループで作成された配列を反復処理します。

ただし、最初のループで作成された配列には、[分割ループパターンを使用してドキュメントから値を読み込む](#reading-values-from-the-document-with-the-split-loop-pattern)セクションの最初のループと同じように、Office オブジェクトのみが含まれているわけでは*ありません*。 これは、Word の Range オブジェクトの処理に必要な情報の一部は、Range オブジェクト自体ではなく、 `jobMapping`配列から取得されるためです。 

そのため、最初のループで作成された配列内のオブジェクトは、2つのプロパティを持つカスタムオブジェクトです。 1つ目は、特定の役職 (つまり、プレースホルダー文字列) に一致する単語範囲の配列で、2番目の配列は、ジョブに割り当てられた人物の名前を提供する文字列です。 これにより、指定された範囲を処理するために必要なすべての情報が範囲を含む同じカスタムオブジェクトに格納されるため、最終ループが簡単に書き込み可能になり、読みやすくなります。 CorrelatedObject を置き換える名前。 _ **correlatedObject**rangesMatchingJob [j]_ は、同じオブジェクトのもう1つのプロパティです。 _ **correlatedObject**: personAssignedToJob_。 

このような分割ループパターンは、このバリエーションによって関連付けられた**オブジェクト**パターンを呼び出します。 一般的な考え方は、最初のループでカスタムオブジェクトの配列が作成されるということです。 各オブジェクトには、Office コレクションオブジェクト内の項目のいずれか (またはそのような項目の配列) の値を持つプロパティがあります。 カスタムオブジェクトには、最後のループで Office オブジェクトを処理するために必要な情報が含まれている他のプロパティがあります。 カスタム関連付けオブジェクトに3つ以上のプロパティがある場合のリンクについては、「[その他のパターンの例](#other-examples-of-these-patterns)」を参照してください。

さらに注意してください。カスタムの関連付けが必要なオブジェクトの配列を作成するためだけに複数のループが発生する場合があります。 これは、ある Office コレクションオブジェクトの各メンバーのプロパティを読み込んで、別のコレクションオブジェクトの処理に使用される情報を収集する必要がある場合に発生することがあります。 (たとえば、アドインでは、列のタイトルに基づいて一部の列のセルに数値の書式を適用するため、コードでは、Excel のテーブル内のすべての列のタイトルを読み取る必要があります)。ただし、ループではなく`context.sync`、ループ間では常に s を続けることができます。 [このようなパターンの](#other-examples-of-these-patterns)例については、「その他の例」を参照してください。

## <a name="other-examples-of-these-patterns"></a>これらのパターンのその他の例

- ループを使用`Array.forEach`する Excel の非常に簡単な例については、「このスタックオーバーフローに対する応答の受け入れ」を参照してください。複数の[コンテキストをキューに追加することができます。コンテキストを同期する前に読み込みます](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)。
- ループを`Array.forEach`使用し、構文を使用`async` / `await`しない Word の簡単な例については、「このスタックオーバーフローの回答」を参照してください。 [Office JavaScript API を使用したコンテンツコントロールを含むすべての段落](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)に対して反復処理を行います。
- TypeScript で記述されている Word の例については、サンプル[Word アドインの Angular2 スタイルチェッカー](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)(特に、ファイルの[word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)) を参照してください。 このメソッドは`for` 、and `Array.forEach`ループが混在しています。
- 高度な Word サンプルの場合は、[この gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)を[スクリプトラボツール](../overview/explore-with-script-lab.md)にインポートします。 Gist を使用するコンテキストについては、「[テキストの置換後に同期されていない](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)スタックオーバーフローの質問」ドキュメントへの受け入れられた応答を参照してください。 この例では、3つのプロパティを持つ、関連付けができるカスタムオブジェクトの種類を作成します。 合計3つのループを使用して、相関オブジェクトの配列を作成し、さらに2つのループを使用して最終的な処理を行います。 `for`とループが混在して`Array.forEach`います。
- スプリットループや相関オブジェクトのパターンの例は厳密ではありませんが、セルの値のセットを1つ`context.sync`だけの他の通貨に変換する方法を示す高度な Excel サンプルがあります。 これを実行するには、[スクリプトラボツール](../overview/explore-with-script-lab.md)を開き、**通貨コンバーター**のサンプルに移動します。 

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>この記事のパターンを使用し*ない*場合は、どうすればよいですか。

Excel は、指定された通話で 5 MB を超えるデータ`context.sync`を読み取ることができません。 この制限を超えると、エラーがスローされます。 (詳細については、「 [Excel データ転送の制限](../develop/common-coding-issues.md#excel-data-transfer-limits)」を参照してください)。この制限が適用されることは非常にまれですが、これがアドインで発生する可能性がある場合は、コードですべてのデータを1つのループ*にロードし*て、ループに従う必要`context.sync`があります。 しかし、 `context.sync`コレクションオブジェクトに対するループの繰り返しが発生しないようにする必要があります。 代わりに、ループ間にを`context.sync`使用して、コレクション内の項目のサブセットを定義し、各サブセットを順番にループします。 これは、サブセットを反復処理する外側のループを使用して`context.sync`構造化し、これらの外側の反復のそれぞれにを含めることができます。
