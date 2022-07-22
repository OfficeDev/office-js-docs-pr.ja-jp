---
title: ループで context.sync メソッドを使用しないでください
description: ループ内で context.sync を呼び出さないようにするために、分割ループパターンと相関オブジェクト パターンを使用する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b0239e05a597949160afbb2604143f3d6626462
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958700"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>ループで context.sync メソッドを使用しないでください

> [!NOTE]
> この記事では、バッチ システムを使用して Office ドキュメントを操作する Excel、Word、OneNote、Visio&mdash;用の 4 つのアプリケーション固有の Office JavaScript API の&mdash;少なくとも 1 つを使用する作業の開始段階を超えているものとします。 特に、呼び出しの `context.sync` 実行内容を把握し、コレクション オブジェクトが何であるかを把握する必要があります。 その段階でない場合は、 [Office JavaScript API の概要](../develop/understanding-the-javascript-api-for-office.md) と、その記事の「アプリケーション固有」の下にリンクされているドキュメントから始めてください。

アプリケーション固有の API モデルの 1 つ (Excel、Word、OneNote、Visio 用) を使用する Office アドインの一部のプログラミング シナリオでは、コレクション オブジェクトのすべてのメンバーから一部のプロパティを読み取り、書き込み、または処理する必要があります。 たとえば、特定のテーブル列のすべてのセルの値を取得する必要がある Excel アドインや、文書内の文字列のすべてのインスタンスを強調表示する必要がある Word アドインなどです。 コレクション オブジェクトのプロパティ内の `items` メンバーを反復処理する必要がありますが、パフォーマンス上の理由から、ループのすべての反復で呼び出 `context.sync` されないようにする必要があります。 すべての呼び出し `context.sync` は、アドインから Office ドキュメントへのラウンド トリップです。 ラウンド トリップを繰り返すとパフォーマンスが低下します。特に、アドインがOffice on the webで実行されている場合は、ラウンド トリップがインターネット経由で行われるためです。

> [!NOTE]
> この記事のすべての例ではループが使用 `for` されていますが、説明されているプラクティスは、次のような配列を反復処理できるループ ステートメントに適用されます。
>
> - `for`
> - `for of`
> - `while`
> - `do while`
>
> また、次のように、関数が渡され、配列内の項目に適用される配列メソッドにも適用されます。
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

最も単純なケースでは、コレクション オブジェクトのメンバーにのみ書き込み、そのプロパティを読み取りません。 たとえば、次のコードは、Word 文書内の "the" のすべてのインスタンスを黄色で強調表示しています。

> [!NOTE]
> 通常は、アプリケーション`run`関数の終了 "}" 文字の直前に最終的`context.sync`な文字 (など`Excel.run``Word.run`) を配置することをお勧めします。 これは、まだ同期されていないキューに `run` 入っているコマンドがある場合にのみ、関数が最後の操作として非表示の呼び出し `context.sync` を行うためです。 この呼び出しが非表示になっているという事実は混乱する可能性があるため、通常は明示的 `context.sync`な呼び出しを追加することをお勧めします。 ただし、この記事では、呼び出し `context.sync`を最小限に抑えることについて説明しているため、完全に不要な最終処理 `context.sync`を追加する方が混乱します。 そのため、この記事では `run`、.

```javascript
await Word.run(async function (context) {
  let startTime, endTime;
  const docBody = context.document.body;

  // search() returns an array of Ranges.
  const searchResults = docBody.search('the', { matchWholeWord: true });
  searchResults.load('font');
  await context.sync();

  // Record the system time.
  startTime = performance.now();

  for (let i = 0; i < searchResults.items.length; i++) {
    searchResults.items[i].font.highlightColor = '#FFFF00';

    await context.sync(); // SYNCHRONIZE IN EACH ITERATION
  }
  
  // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

  // Record the system time again then calculate how long the operation took.
  endTime = performance.now();
  console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
})
```

上記のコードは、Windows 上の Word で 200 個の "the" インスタンスを含むドキュメントで完了するまでに 1 秒を要しました。 ただし、ループ内の `await context.sync();` 行がコメントアウトされ、ループのコメントが解除された直後の同じ行の場合、操作にかかった時間は 1 秒の 1/10 だけです。 Web 上の Word (ブラウザーとして Edge を使用) では、ループ内の同期に 3 秒かかり、ループ後の同期では 1 秒の 6/10 秒だけで、約 5 倍の速さでした。 "the" のインスタンスが 2,000 個のドキュメントでは、ループ内の同期で (Word on the web で) 80 秒かかり、ループ後の同期では約 4 秒かかり、約 20 倍の速さでした。

> [!NOTE]
> 同期が同時に実行された場合に、同期内部のバージョンが高速に実行されるかどうかを確認することをお勧めします。これは、キーワードを `await` 前 `context.sync()`から削除するだけで実行できます。 これにより、ランタイムは同期を開始し、同期が完了するのを待たずにすぐにループの次のイテレーションを開始します。 ただし、これは、このような理由からループから完全に外 `context.sync` れるほど優れたソリューションではありません。
>
> - 同期バッチ ジョブのコマンドがキューに登録されているのと同様に、バッチ ジョブ自体は Office でキューに登録されますが、Office ではキュー内の 50 個以下のバッチ ジョブがサポートされます。 それ以上、エラーがトリガーされます。 そのため、ループ内に 50 回以上の反復がある場合は、キュー サイズを超える可能性があります。 反復回数が多いほど、これが発生する可能性が高くなります。
> - "同時" は同時に意味しません。 複数の同期操作を実行するには、1 つを実行するよりも時間がかかります。
> - 同時操作は、開始したのと同じ順序で完了することは保証されません。 前の例では、"the" という単語が強調表示される順序は関係ありませんが、コレクション内の項目を順番に処理することが重要なシナリオがあります。

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>分割ループ パターンを使用してドキュメントから値を読み取る

ループ内で s を `context.sync`回避することは、コードがコレクション項目のプロパティを *読み取る* 必要があるときに、それぞれが処理されるときに、より困難になります。 コードで Word 文書内のすべてのコンテンツ コントロールを反復処理し、各コントロールに関連付けられている最初の段落のテキストを記録する必要があるとします。 プログラミングの本能によって、コントロールをループし、各 (最初の) 段落のプロパティを `text` 読み込み、ドキュメントのテキストをプロキシ段落オブジェクトに設定してログに記録する呼び出 `context.sync` しが発生する可能性があります。 次に例を示します。

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

このシナリオでは、ループ内のパターンを `context.sync` 回避するために、 **分割** ループ パターンを呼び出すパターンを使用する必要があります。 パターンの具体的な例を見てから、その正式な説明を見てみましょう。 前のコード スニペットに分割ループ パターンを適用する方法を次に示します。 このコードについては、次の点に注意してください。

- これで 2 つのループがあり、 `context.sync` その間にループが存在するため、どちらのループ内にもありません `context.sync` 。
- 最初のループは、コレクション オブジェクト内の項目を反復処理し、元の`text`ループと同じようにプロパティを読み込みますが、最初のループでは、プロキシ オブジェクトのプロパティを設定`text`する要素`context.sync`が含まれなくなったため、段落テキストを`paragraph`記録できません。 代わりに、オブジェクトを `paragraph` 配列に追加します。
- 2 番目のループは、最初のループによって作成された配列を反復処理し、各項目`paragraph`のログを`text`記録します。 これは、2 つのループの `context.sync` 間に含まれるすべてのプロパティが設定されたために発生する `text` 可能性があります。

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

上の例では、a を含む `context.sync` ループを分割ループ パターンに変換するための次の手順を示します。

1. ループを 2 つのループに置き換えます。
2. コレクションを反復処理し、各項目を配列に追加する最初のループを作成し、コードで読み取る必要がある項目のプロパティも読み込みます。
3. 最初のループに続いて、読み込まれたプロパティをプロキシ オブジェクトに設定する呼び出し `context.sync` 。
4. 2 番目の `context.sync` ループに従って、最初のループで作成された配列を反復処理し、読み込まれたプロパティを読み取ります。

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>相関オブジェクト パターンを使用してドキュメント内のオブジェクトを処理する

コレクション内の項目を処理するには、アイテム自体にないデータが必要になる、より複雑なシナリオを考えてみましょう。 このシナリオでは、テンプレートから作成された文書を定型句で操作する Word アドインを想定しています。 テキストに散在しているのは、"{コーディネーター}"、"{Deputy}"、"{Manager}" というプレースホルダー文字列の 1 つ以上のインスタンスです。 アドインは、各プレースホルダーを一部のユーザーの名前に置き換えます。 この記事では、アドインの UI は重要ではありません。 たとえば、3 つのテキスト ボックスを含む作業ウィンドウがあり、それぞれがプレースホルダーの 1 つでラベル付けされます。 ユーザーが各テキスト ボックスに名前を入力し、 **置換** ボタンを押します。 ボタンのハンドラーは、プレースホルダーに名前をマップし、各プレースホルダーを割り当てられた名前に置き換える配列を作成します。

この UI を使用して実際にアドインを生成してコードを試す必要はありません。 [Script Lab ツール](../overview/explore-with-script-lab.md)を使用して、重要なコードのプロトタイプを作成できます。 次の代入ステートメントを使用して、マッピング配列を作成します。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

次のコードは、ループ内で使用 `context.sync` した場合に、各プレースホルダーを割り当てられた名前に置き換える方法を示しています。

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

上記のコードでは、外部ループと内部ループがあります。 それぞれには `context.sync`、 . この記事の最初のコード スニペットに基づいて、 `context.sync` 内側のループ内を内側のループの後に移動できることがわかります。 しかし、それでもコード `context.sync` は外側のループに (実際には 2 つ) 残されます。 次のコードは、ループから削除 `context.sync` する方法を示しています。 以下のコードについて説明します。

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

このコードでは、分割ループ パターンが使用されることに注意してください。

- 前の例の外側のループは 2 つに分割されています。 (2 番目のループには内部ループがあります。これは、コードが一連のジョブ (またはプレースホルダー) を反復処理し、そのセット内で一致する範囲を反復処理するためです)。
- 各メジャー ループの `context.sync` 後は存在しますが、ループ内にはありません `context.sync` 。
- 2 番目のメジャー ループは、最初のループで作成された配列を反復処理します。

ただし、最初のループで作成された配列には、最初のループが分割ループ パターンを使用して [ドキュメントから値を読み取](#read-values-from-the-document-with-the-split-loop-pattern)るセクションで行ったように、Office オブジェクトのみが含 *まれていません*。 これは、Word Range オブジェクトを処理するために必要な情報の一部が Range オブジェクト自体ではなく、配列から `jobMapping` 取得されるためです。

したがって、最初のループで作成された配列内のオブジェクトは、2 つのプロパティを持つカスタム オブジェクトです。 1 つ目は、特定の役職 (つまりプレースホルダー文字列) と一致する Word 範囲の配列であり、2 つ目は、ジョブに割り当てられたユーザーの名前を提供する文字列です。 これにより、特定の範囲を処理するために必要なすべての情報が、その範囲を含む同じカスタム オブジェクトに含まれているため、最終的なループは書き込みが簡単で読みやすくなります。 _correlatedObject.rangesMatchingJob.items[j] を_ 置き換える必要がある名前は、同じオブジェクトのもう 1 つのプロパティである _**、correlatedObject.personAssignedToJob** です_。

この分割ループ パターンのバリエーションを **、相関オブジェクト** パターンと呼びます。 一般的な考え方は、最初のループによってカスタム オブジェクトの配列が作成されるということです。 各オブジェクトには、値が Office コレクション オブジェクト内の項目の 1 つ (またはそのような項目の配列) であるプロパティがあります。 カスタム オブジェクトには他のプロパティがあり、それぞれが最後のループで Office オブジェクトを処理するために必要な情報を提供します。 カスタム相関オブジェクトに 2 つ以上のプロパティがある例へのリンクについては、 [これらのパターンのその他の例](#other-examples-of-these-patterns) を参照してください。

もう 1 つの注意点: カスタム相関オブジェクトの配列を作成するだけで、複数のループが必要な場合があります。 これは、1 つの Office コレクション オブジェクトの各メンバーのプロパティを読み取って、別のコレクション オブジェクトの処理に使用される情報を収集する必要がある場合に発生する可能性があります。 (たとえば、アドインがその列のタイトルに基づいて一部の列のセルに数値形式を適用するため、コードでは Excel テーブル内のすべての列のタイトルを読み取る必要があります)。ただし、ループではなく、ループ間の s は常に保持 `context.sync`できます。 例については、 [これらのパターンのその他の例](#other-examples-of-these-patterns) のセクションを参照してください。

## <a name="other-examples-of-these-patterns"></a>これらのパターンのその他の例

- ループを使用 `Array.forEach` する Excel の非常に簡単な例については、この Stack Overflow の質問に対する受け入れられた回答を参照してください。 [context.sync の前に複数の context.load をキューに入れることは可能ですか?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- ループを使用し、構文を使用しない Word の簡単な例については、Office JavaScript API を使用`Array.forEach`して[コンテンツ コントロールを使用](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)`async`/`await`してすべての段落を反復処理するという、スタック オーバーフローの質問に対する受け入れられた回答を参照してください。
- TypeScript で記述されている Word の例については、サンプル [の Word アドイン Angular2 スタイル チェッカー](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker) (特にファイル [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)) を参照してください。 これには、ループとループの `for` 組み合わせがあります `Array.forEach` 。
- 高度な Word サンプルの場合は、[この要を](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)[Script Lab ツール](../overview/explore-with-script-lab.md)にインポートします。 要点を使用するコンテキストについては、テキストの置換後にスタック オーバーフローの質問 [ドキュメントが同期されない](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)という回答を参照してください。 このサンプルでは、3 つのプロパティを持つカスタムの相関オブジェクト型を作成します。 合計 3 つのループを使用して相関オブジェクトの配列を作成し、さらに 2 つのループを使用して最終的な処理を行います。 ループと`Array.forEach`ループの組み合わせ`for`があります。
- 分割ループまたは相関オブジェクト パターンの例を厳密には示していませんが、セル値のセットを 1 つだけ `context.sync`で他の通貨に変換する方法を示す高度な Excel サンプルがあります。 試すには、[Script Lab ツール](../overview/explore-with-script-lab.md)を開き、**Currency Converter** サンプルに移動します。

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>この記事のパターンをいつ使用 *しない* 必要がありますか?

Excel では、特定の呼び出し `context.sync`で 5 MB を超えるデータを読み取ることはできません。 この制限を超えると、エラーがスローされます。 (詳細については、Office アドインの [リソース制限とパフォーマンスの最適化](resource-limits-and-performance-optimization.md#excel-add-ins)の「Excel アドイン」セクションを参照してください)。この制限に近づくことは非常にまれですが、アドインでこれが発生する可能性がある場合は、コードですべてのデータを 1 つのループに読み込み、ループ`context.sync`に従う *必要はありません*。 ただし、コレクション オブジェクトに対する `context.sync` ループの繰り返しごとに発生しないようにする必要があります。 代わりに、コレクション内の項目のサブセットを定義し、各サブセットをループ間で `context.sync` 順番にループします。 これを構成するには、サブセットを反復処理し、これらの各外部反復を `context.sync` 含む外部ループを使用します。
