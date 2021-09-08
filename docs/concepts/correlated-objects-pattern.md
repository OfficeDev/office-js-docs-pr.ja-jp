---
title: ループで context.sync メソッドを使用しないでください
description: ループ内での context.sync の呼び出しを回避するために、分割ループと相関オブジェクト パターンを使用する方法について説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 85230378f40be06c7f3385f5dde88ecaba503cb5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937709"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>ループで context.sync メソッドを使用しないでください

> [!NOTE]
> この記事では、バッチ システムを使用して Office ドキュメントを操作する &mdash; Excel、Word、OneNote、および Visio の 4 つのアプリケーション固有の Office JavaScript API の少なくとも 1 つを操作する最初の段階を超えていると仮定します。 &mdash; 特に、呼び出しが何を行うのかを知り、 `context.sync` コレクション オブジェクトが何かを知る必要があります。 その段階ではない場合は、まず[JavaScript API](../develop/understanding-the-javascript-api-for-office.md)の Officeと、その記事の 「アプリケーション固有」の下にリンクされているドキュメントについてを参照してください。

Office アドインで、アプリケーション固有の API モデル (Excel、Word、OneNote、および Visio 用) のいずれかを使用する一部のプログラミング シナリオでは、コレクション オブジェクトのすべてのメンバーからいくつかのプロパティを読み取り、書き込み、または処理する必要があります。 たとえば、特定のテーブル列内のすべてのセルの値を取得する必要がある Excel アドインや、ドキュメント内の文字列のすべてのインスタンスを強調表示する必要がある Word アドインなどです。 コレクション オブジェクトのプロパティ内のメンバーを反復処理する必要がありますが、パフォーマンス上の理由から、ループのすべての反復で呼び出しを避 `items` `context.sync` ける必要があります。 すべての呼び出しは、アドインからドキュメントへの `context.sync` ラウンド トリップOfficeです。 ラウンド トリップを繰り返す場合はパフォーマンスが低下します。特に、ラウンド トリップがインターネットを通Office on the webで実行されている場合は特にパフォーマンスが低下します。

> [!NOTE]
> この記事のすべての例ではループを使用しますが、説明するプラクティスは、次のような配列を反復処理できるループ ステートメント `for` に適用されます。
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> また、関数が渡され、次のような配列内のアイテムに適用される配列メソッドにも適用されます。
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

最も単純なケースでは、コレクション オブジェクトのメンバーにのみ書き込み、プロパティを読み取る必要があります。 たとえば、次のコードでは、Word ドキュメントの "the" のすべてのインスタンスが黄色で強調表示されます。

> [!NOTE]
> 一般に、アプリケーション メソッドの終了 "}" 文字の直前に final を設定する (、など) の方が `context.sync` `run` `Excel.run` `Word.run` 良い方法です。 これは、メソッドが最後の呼び出しとして非表示の呼び出しを行い、まだ同期されていないキューに入っているコマンドがある場合にのみ `run` `context.sync` 行うためです。 この呼び出しが非表示であるという事実はわかりにくい場合があります。そのため、通常は明示的に追加することをお勧めします `context.sync` 。 ただし、この記事では呼び出しを最小限に抑えることについて考えると、実際には完全に不要な最終項目を追加する方が `context.sync` 複雑です `context.sync` 。 したがって、この記事では、同期されていないコマンドがない場合は、このコマンドを残します `run` 。

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

前のコードは、Word on Windows で 200 インスタンスの "the" を含むドキュメントで完了するために 1 秒Windows。 ただし、ループ内の行がコメントアウトされ、ループがコメント解除された直後に同じ行が返された場合、操作は `await context.sync();` 1/10 分の 1 秒しかかからなかった。 Word on the web (ブラウザーとして Edge を使用) では、ループ内の同期に 3 秒かかり、ループ後の同期では 6/10 分の 1 秒で、約 5 倍速くなります。 2000 インスタンスの "the" を含むドキュメントでは、ループ内の同期に Word on the web 80 秒かかり 、ループの同期に 4 秒しかかからなかり、約 20 倍速くなります。

> [!NOTE]
> 同期が同時に実行された場合に同期インサイド ザ ループ バージョンの実行速度が速くなるかどうかを確認する必要があります。これは、キーワードを前から削除するだけで実行できます `await` `context.sync()` 。 これにより、ランタイムは同期を開始し、同期が完了するのを待たずに、ループの次の反復を直ちに開始します。 ただし、このような理由からループから完全に移動するほど良いソリューション `context.sync` ではありません。
>
> - 同期バッチ ジョブのコマンドがキューに入れられますが、バッチ ジョブ自体は Office でキューに入れられますが、Office はキュー内で 50 以下のバッチ ジョブをサポートします。 それ以上のトリガー エラー。 したがって、ループ内に 50 回を超える反復がある場合は、キュー サイズを超える可能性があります。 繰り返し回数が多い場合は、このようなことが起こる可能性が高い。 
> - "同時に" とは、同時に意味する意味ではありません。 複数の同期操作を実行するには、1 つを実行するよりも時間がかかります。
> - 同時操作は、開始した順序と同じ順序で完了するとは保証されません。 前の例では、"the" という単語が強調表示される順序は関係ありませんが、コレクション内のアイテムを順番に処理することが重要なシナリオがあります。

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>分割ループ パターンを使用してドキュメントから値を読み取る

ループ内の s を避けることは、コードがコレクション アイテムのプロパティを読み取る必要があるときに、各コレクション アイテムを処理するときに `context.sync` 、より困難になります。  コードで Word ドキュメント内のすべてのコンテンツ コントロールを反復処理し、各コントロールに関連付けられた最初の段落のテキストを記録する必要があるものとします。 プログラミングの本能により、コントロールをループ処理し、各 `text` (最初の) 段落のプロパティを読み込み、プロキシ段落オブジェクトにドキュメントのテキストを設定する呼び出しを行い、ログに記録する場合があります。 `context.sync` 次に例を示します。

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

このシナリオでは、ループ内にループが含まれるのを避けるために、分割ループ パターンを呼び出 `context.sync` すパターン **を使用する必要** があります。 パターンの具体的な例を見てから、そのパターンの正式な説明を確認します。 前のコード スニペットに分割ループ パターンを適用する方法を次に示します。 このコードについては、次の点に注意してください。

- 2 つのループが作成され、その間にループが生じ、どちらのループ `context.sync` `context.sync` も内部にありません。
- 最初のループは、コレクション オブジェクト内のアイテムを反復処理し、元のループと同じ方法でプロパティを読み込むが、プロキシ オブジェクトのプロパティを設定する a が含まれるので、最初のループでは段落のテキストをログに記録できません。 `text` `context.sync` `text` `paragraph` 代わりに、オブジェクトを `paragraph` 配列に追加します。
- 2 番目のループは、最初のループによって作成された配列を反復処理し、各アイテム `text` のログを記録 `paragraph` します。 これは、2 つのループ `context.sync` の間に含まれるすべてのプロパティが設定されたためです `text` 。

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

前の例では、a を含むループを分割ループ パターンに変換する手順 `context.sync` を示します。

1. ループを 2 つのループに置き換える。
2. コレクションを反復処理し、各アイテムを配列に追加し、コードで読み取る必要があるアイテムのプロパティも読み込む最初のループを作成します。
3. 最初のループに続き、プロキシ `context.sync` オブジェクトに読み込まれたプロパティを設定する呼び出しを行います。
4. 2 番目のループに従って、最初のループで作成された配列を反復処理し、読み込まれた `context.sync` プロパティを読み取ります。

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>関連付けオブジェクト パターンを使用してドキュメント内のオブジェクトを処理する

コレクション内のアイテムを処理するには、アイテム自体に含されていないデータが必要な、より複雑なシナリオについて考えます。 このシナリオでは、テンプレートから作成されたいくつかの定型文を含むドキュメントを操作する Word アドインを想定しています。 テキストに散在するインスタンスは、"{Coordinator}"、"{Deputy}"、および "{Manager}" というプレースホルダー文字列の 1 つ以上のインスタンスです。 アドインは、各プレースホルダーを一部のユーザーの名前に置き換える。 この記事では、アドインの UI は重要ではありません。 たとえば、作業ウィンドウに 3 つのテキスト ボックスが表示され、それぞれにプレースホルダーの 1 つが付きます。 ユーザーは、各テキスト ボックスに名前を入力し、[置換] ボタンを **押** します。 ボタンのハンドラーは、名前をプレースホルダーにマップする配列を作成し、各プレースホルダーを割り当てられた名前に置き換える。

この UI を使用して実際にアドインを作成して、コードを試す必要はありません。 このツールを使用[Script Labコード](../overview/explore-with-script-lab.md)のプロトタイプを作成できます。 マッピング配列を作成するには、次の代入ステートメントを使用します。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

次のコードは、ループ内で使用した場合に、各プレースホルダーを割り当てられた名前に置き換える方法 `context.sync` を示しています。

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

前のコードでは、外部ループと内部ループがあります。 これらの各ファイルには、 `context.sync` が含まれる。 この記事の最初のコード スニペットに基づいて、内部ループ内を内部ループの後に移動できる `context.sync` 可能性があります。 しかし、それでもコードは外側のループに (実際には 2 `context.sync` つ) 残ります。 次のコードは、ループから削除する `context.sync` 方法を示しています。 以下のコードについて説明します。

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

コードでは、分割ループ パターンが使用されます。

- 前の例の外側のループは 2 つに分割されています。 (2 番目のループには内部ループがあります。これは、コードが一連のジョブ (またはプレースホルダー) を反復処理し、そのセット内で一致する範囲を反復処理している場合に予期されます)。
- 各メジャー ループ `context.sync` の後に発生しますが、ループ `context.sync` 内には含めはありません。
- 2 番目のメジャー ループは、最初のループで作成された配列を反復処理します。

ただし、最初のループで作成された配列には、セクション「分割ループ パターンを使用してドキュメントから値を読み取る」セクションで行った最初のループと同様に、Office オブジェクト[だけが含まれます](#read-values-from-the-document-with-the-split-loop-pattern)。 これは、Word Range オブジェクトの処理に必要な情報の一部が Range オブジェクト自体ではなく、配列に含まれるため `jobMapping` です。

したがって、最初のループで作成された配列内のオブジェクトは、2 つのプロパティを持つカスタム オブジェクトです。 1 つ目は、特定の役職 (プレースホルダー文字列) に一致する Word の範囲の配列で、2 つ目はジョブに割り当てられたユーザーの名前を示す文字列です。 これにより、指定した範囲を処理するために必要なすべての情報が、その範囲を含む同じカスタム オブジェクトに含まれているため、最終的なループを簡単に記述し、読みやすくします。 _**correlatedObject**.rangesMatchingJob.items[j]_ を置き換える必要がある名前は、同じオブジェクトのもう 1 つのプロパティです _**。correlatedObject**.personAssignedToJob_ です。

この分割ループ パターンのバリエーションを、相関オブジェクト **パターンと呼** ぶ。 一般的な考え方は、最初のループがカスタム オブジェクトの配列を作成する方法です。 各オブジェクトには、コレクション オブジェクト内の項目の 1 つ (またはOfficeの配列) の値を持つプロパティがあります。 カスタム オブジェクトには他のプロパティが含まれています。各プロパティには、最終ループ内のオブジェクトを処理するためにOffice情報が提供されます。 カスタム関連付 [けオブジェクト](#other-examples-of-these-patterns) に複数のプロパティがある例へのリンクについては、「これらのパターンのその他の例」を参照してください。

もう 1 つの注意点: カスタム相関オブジェクトの配列を作成するために複数のループが必要な場合があります。 これは、別のコレクション オブジェクトの処理に使用される情報を収集するために、Office コレクション オブジェクトの各メンバーのプロパティを読み取る必要がある場合に発生します。 (たとえば、アドインは、その列のタイトルに基づいていくつかの列のセルに数値形式を適用するつもりなので、Excel テーブル内のすべての列のタイトルを読み取る必要があります)。ただし、ループではなく、ループの間に s を常 `context.sync` に保持できます。 例については、 [セクション「これらのパターンのその他の例](#other-examples-of-these-patterns) 」を参照してください。

## <a name="other-examples-of-these-patterns"></a>これらのパターンの他の例

- ループを使用する Excel の非常に簡単な例については、このスタック オーバーフローの質問に対する受け入れ可能な回答を参照してください。context.sync の前に複数の `Array.forEach` [context.load](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)をキューに入れられますか?
- ループを使用し、構文を使用しない Word の簡単な例については、このスタック オーバーフローの質問に対する受け入れ可能な回答を参照してください `Array.forEach` `async` / `await` [。Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)を使用してコンテンツ コントロールを使用してすべての段落を反復処理します。
- TypeScript で記述されている Word の例については、サンプルの Word アドイン [Angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)スタイル チェッカー (特に [ ument.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)のファイルword.docを参照してください。 これは、ループの混合 `for` を `Array.forEach` 持っています。
- 高度な Word サンプルの場合は、この[gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)を新しいツール[にScript Labします](../overview/explore-with-script-lab.md)。 gist を使用するコンテキストについては、テキストの置換後に同期されていないスタック オーバーフローの質問ドキュメントに対する受け入れ可能 [な回答を参照してください](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)。 このサンプルでは、3 つのプロパティを持つカスタムの関連付けオブジェクトの種類を作成します。 合計 3 つのループを使用して相関オブジェクトの配列を作成し、さらに 2 つのループを使用して最終的な処理を実行します。 ループとループの混合 `for` `Array.forEach` があります。
- 分割ループや相関オブジェクト パターンの厳密な例ではありませんが、セル値のセットを単一の通貨で他の通貨に変換する方法を示す高度な Excel サンプルがあります `context.sync` 。 このツールを試す場合は、Script Lab [ツールを](../overview/explore-with-script-lab.md)開き **、Currency Converter サンプルに移動** します。

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>いつこの *記事の* パターンを使う必要がありますか?

Excel呼び出しでは、5 MB を超えるデータを読み取る必要があります `context.sync` 。 この制限を超えると、エラーがスローされます。 (詳細については、「Excel アドインのリソース制限とパフォーマンスの最適化」[](resource-limits-and-performance-optimization.md#excel-add-ins)の「Office アドイン」を参照してください。この制限に近づくことは非常にまれですが、アドインでこれが発生する可能性がある場合は、コードですべてのデータを 1 つのループで読み込み、ループに従う必要はありません `context.sync` 。 ただし、コレクション オブジェクトに対するループの繰り返しごとに a `context.sync` を使用しないようにする必要があります。 代わりに、コレクション内のアイテムのサブセットを定義し、ループ間で各サブセットを順番 `context.sync` にループします。 これは、サブセットを反復処理し、これらの外側の各反復に含まれる外部ループ `context.sync` で構成できます。
