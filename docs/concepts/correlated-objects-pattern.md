---
title: ループで context.sync メソッドを使用しないでください
description: ループ内での context.sync の呼び出しを回避するために、分割ループと相関オブジェクト パターンを使用する方法について説明します。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 64cfd5cd350746ba07e1a98986a4bd7811431475
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349141"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a><span data-ttu-id="177ab-103">ループで context.sync メソッドを使用しないでください</span><span class="sxs-lookup"><span data-stu-id="177ab-103">Avoid using the context.sync method in loops</span></span>

> [!NOTE]
> <span data-ttu-id="177ab-104">この記事では、バッチ システムを使用して Office ドキュメントを操作する &mdash; Excel、Word、OneNote、および Visio の 4 つのアプリケーション固有の Office JavaScript API の少なくとも 1 つを操作する最初の段階を超えていると仮定します。 &mdash;</span><span class="sxs-lookup"><span data-stu-id="177ab-104">This article assumes that you're beyond the beginning stage of working with at least one of the four application-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document.</span></span> <span data-ttu-id="177ab-105">特に、呼び出しが何を行うのかを知り、 `context.sync` コレクション オブジェクトが何かを知る必要があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-105">In particular, you should know what a call of `context.sync` does and you should know what a collection object is.</span></span> <span data-ttu-id="177ab-106">その段階ではない場合は、まず[JavaScript API](../develop/understanding-the-javascript-api-for-office.md)の Officeと、その記事の 「アプリケーション固有」の下にリンクされているドキュメントについてを参照してください。</span><span class="sxs-lookup"><span data-stu-id="177ab-106">If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "application-specific" in that article.</span></span>

<span data-ttu-id="177ab-107">Office アドインで、アプリケーション固有の API モデル (Excel、Word、OneNote、および Visio 用) のいずれかを使用する一部のプログラミング シナリオでは、コレクション オブジェクトのすべてのメンバーからいくつかのプロパティを読み取り、書き込み、または処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-107">For some programming scenarios in Office Add-ins that use one of the application-specific API models (for Excel, Word, OneNote, and Visio), your code needs to read, write, or process some property from every member of a collection object.</span></span> <span data-ttu-id="177ab-108">たとえば、特定のテーブル列内のすべてのセルの値を取得する必要がある Excel アドインや、ドキュメント内の文字列のすべてのインスタンスを強調表示する必要がある Word アドインなどです。</span><span class="sxs-lookup"><span data-stu-id="177ab-108">For example, an Excel add-in that needs to get the values of every cell in a particular table column or a Word add-in that needs to highlight every instance of a string in the document.</span></span> <span data-ttu-id="177ab-109">コレクション オブジェクトのプロパティ内のメンバーを反復処理する必要がありますが、パフォーマンス上の理由から、ループのすべての反復で呼び出しを避 `items` `context.sync` ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-109">You need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you need to avoid calling `context.sync` in every iteration of the loop.</span></span> <span data-ttu-id="177ab-110">すべての呼び出しは、アドインからドキュメントへの `context.sync` ラウンド トリップOfficeです。</span><span class="sxs-lookup"><span data-stu-id="177ab-110">Every call of `context.sync` is a round trip from the add-in to the Office document.</span></span> <span data-ttu-id="177ab-111">ラウンド トリップを繰り返す場合はパフォーマンスが低下します。特に、ラウンド トリップがインターネットを通Office on the webで実行されている場合は特にパフォーマンスが低下します。</span><span class="sxs-lookup"><span data-stu-id="177ab-111">Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.</span></span>

> [!NOTE]
> <span data-ttu-id="177ab-112">この記事のすべての例ではループを使用しますが、説明するプラクティスは、次のような配列を反復処理できるループ ステートメント `for` に適用されます。</span><span class="sxs-lookup"><span data-stu-id="177ab-112">All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:</span></span>
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> <span data-ttu-id="177ab-113">また、関数が渡され、次のような配列内のアイテムに適用される配列メソッドにも適用されます。</span><span class="sxs-lookup"><span data-stu-id="177ab-113">They also apply to any array method to which a function is passed and applied to the items in the array, including the following:</span></span>
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

## <a name="writing-to-the-document"></a><span data-ttu-id="177ab-114">ドキュメントへの書き込み</span><span class="sxs-lookup"><span data-stu-id="177ab-114">Writing to the document</span></span>

<span data-ttu-id="177ab-115">最も単純なケースでは、コレクション オブジェクトのメンバーにのみ書き込み、プロパティを読み取る必要があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-115">In the simplest case, you are only writing to members of a collection object, not reading their properties.</span></span> <span data-ttu-id="177ab-116">たとえば、次のコードでは、Word ドキュメントの "the" のすべてのインスタンスが黄色で強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="177ab-116">For example, the following code highlights in yellow every instance of "the" in a Word document.</span></span>

> [!NOTE]
> <span data-ttu-id="177ab-117">一般に、アプリケーション メソッドの終了 "}" 文字の直前に最終値 (、、など) を付けるのが `context.sync` `run` `Excel.run` `Word.run` 良い方法です。</span><span class="sxs-lookup"><span data-stu-id="177ab-117">It is generally a good practice to put have a final `context.sync` just before the closing "}" character of the application `run` method (such as `Excel.run`, `Word.run`, etc.).</span></span> <span data-ttu-id="177ab-118">これは、メソッドが最後の呼び出しとして非表示の呼び出しを行い、まだ同期されていないキューに入っているコマンドがある場合にのみ `run` `context.sync` 行うためです。</span><span class="sxs-lookup"><span data-stu-id="177ab-118">This is because the `run` method makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that have not yet been synchronized.</span></span> <span data-ttu-id="177ab-119">この呼び出しが非表示であるという事実はわかりにくい場合があります。そのため、通常は明示的に追加することをお勧めします `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-119">The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`.</span></span> <span data-ttu-id="177ab-120">ただし、この記事では呼び出しを最小限に抑えることについて考えると、実際には完全に不要な最終項目を追加する方が `context.sync` 複雑です `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-120">However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`.</span></span> <span data-ttu-id="177ab-121">したがって、この記事では、同期されていないコマンドがない場合は、このコマンドを残します `run` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-121">So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.</span></span>

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

<span data-ttu-id="177ab-122">前のコードは、Word on Windows で 200 インスタンスの "the" を含むドキュメントで完了するために 1 秒Windows。</span><span class="sxs-lookup"><span data-stu-id="177ab-122">The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows.</span></span> <span data-ttu-id="177ab-123">ただし、ループ内の行がコメントアウトされ、ループがコメント解除された直後に同じ行が返された場合、操作は `await context.sync();` 1/10 分の 1 秒しかかからなかった。</span><span class="sxs-lookup"><span data-stu-id="177ab-123">But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second.</span></span> <span data-ttu-id="177ab-124">Word on the web (ブラウザーとして Edge を使用) では、ループ内の同期に 3 秒かかり、ループ後の同期では 6/10 分の 1 秒で、約 5 倍速くなります。</span><span class="sxs-lookup"><span data-stu-id="177ab-124">In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster.</span></span> <span data-ttu-id="177ab-125">2000 インスタンスの "the" を含むドキュメントでは、ループ内の同期に Word on the web 80 秒かかり 、ループの同期に 4 秒しかかからなかり、約 20 倍速くなります。</span><span class="sxs-lookup"><span data-stu-id="177ab-125">In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.</span></span>

> [!NOTE]
> <span data-ttu-id="177ab-126">同期が同時に実行された場合に同期インサイド ザ ループ バージョンの実行速度が速くなるかどうかを確認する必要があります。これは、キーワードを前から削除するだけで実行できます `await` `context.sync()` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-126">It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`.</span></span> <span data-ttu-id="177ab-127">これにより、ランタイムは同期を開始し、同期が完了するのを待たずに、ループの次の反復を直ちに開始します。</span><span class="sxs-lookup"><span data-stu-id="177ab-127">This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete.</span></span> <span data-ttu-id="177ab-128">ただし、これは、次の理由でループから完全に移動するほど良 `context.sync` い解決策ではありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-128">However, this is not as good a solution as moving the `context.sync` out of the loop entirely for these reasons:</span></span>
>
> - <span data-ttu-id="177ab-129">同期バッチ ジョブのコマンドがキューに入れられますが、バッチ ジョブ自体は Office でキューに入れられますが、Office はキュー内で 50 以下のバッチ ジョブをサポートします。</span><span class="sxs-lookup"><span data-stu-id="177ab-129">Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue.</span></span> <span data-ttu-id="177ab-130">それ以上のトリガー エラー。</span><span class="sxs-lookup"><span data-stu-id="177ab-130">Any more triggers errors.</span></span> <span data-ttu-id="177ab-131">したがって、ループ内に 50 回を超える反復がある場合は、キュー サイズを超える可能性があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-131">So, if there are more than 50 iterations in a loop, there is a chance that the queue size is exceeded.</span></span> <span data-ttu-id="177ab-132">繰り返し回数が多い場合は、このようなことが起こる可能性が高い。</span><span class="sxs-lookup"><span data-stu-id="177ab-132">The greater the number of iterations, the greater the chance of this happening.</span></span> 
> - <span data-ttu-id="177ab-133">"同時に" とは、同時に意味する意味ではありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-133">"Concurrently" does not mean simultaneously.</span></span> <span data-ttu-id="177ab-134">複数の同期操作を実行するには、1 つを実行するよりも時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="177ab-134">It would still take longer to execute multiple synchronization operations than to execute one.</span></span>
> - <span data-ttu-id="177ab-135">同時操作は、開始した順序と同じ順序で完了するとは保証されません。</span><span class="sxs-lookup"><span data-stu-id="177ab-135">Concurrent operations are not guaranteed to complete in the same order in which they started.</span></span> <span data-ttu-id="177ab-136">前の例では、"the" という単語が強調表示される順序は関係ありませんが、コレクション内のアイテムを順番に処理することが重要なシナリオがあります。</span><span class="sxs-lookup"><span data-stu-id="177ab-136">In the preceding example, it doesn't matter what order the  word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.</span></span>

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a><span data-ttu-id="177ab-137">分割ループ パターンを使用してドキュメントから値を読み取る</span><span class="sxs-lookup"><span data-stu-id="177ab-137">Reading values from the document with the split loop pattern</span></span>

<span data-ttu-id="177ab-138">ループ内の s を避けることは、コードがコレクション アイテムのプロパティを読み取る必要があるときに、各コレクション アイテムを処理するときに `context.sync` 、より困難になります。 </span><span class="sxs-lookup"><span data-stu-id="177ab-138">Avoiding `context.sync`s inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one.</span></span> <span data-ttu-id="177ab-139">コードで Word ドキュメント内のすべてのコンテンツ コントロールを反復処理し、各コントロールに関連付けられた最初の段落のテキストを記録する必要があるものとします。</span><span class="sxs-lookup"><span data-stu-id="177ab-139">Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control.</span></span> <span data-ttu-id="177ab-140">プログラミングの本能により、コントロールをループ処理し、各 `text` (最初の) 段落のプロパティを読み込み、プロキシ段落オブジェクトにドキュメントのテキストを設定する呼び出しを行い、ログに記録する場合があります。 `context.sync`</span><span class="sxs-lookup"><span data-stu-id="177ab-140">Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it.</span></span> <span data-ttu-id="177ab-141">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="177ab-141">The following is an example.</span></span>

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

<span data-ttu-id="177ab-142">このシナリオでは、ループ内にループが含まれるのを避けるために、分割ループ パターンを呼び出 `context.sync` すパターン **を使用する必要** があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-142">In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern.</span></span> <span data-ttu-id="177ab-143">パターンの具体的な例を見てから、そのパターンの正式な説明を確認します。</span><span class="sxs-lookup"><span data-stu-id="177ab-143">Let's see a concrete example of the pattern before we get to a formal description of it.</span></span> <span data-ttu-id="177ab-144">前のコード スニペットに分割ループ パターンを適用する方法を次に示します。</span><span class="sxs-lookup"><span data-stu-id="177ab-144">Here's how the split loop pattern can be applied to the preceding code snippet.</span></span> <span data-ttu-id="177ab-145">このコードについては以下の点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="177ab-145">Note the following about this code.</span></span>

- <span data-ttu-id="177ab-146">2 つのループが作成され、その間にループが生じ、どちらのループ `context.sync` `context.sync` も内部にありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-146">There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.</span></span>
- <span data-ttu-id="177ab-147">最初のループは、コレクション オブジェクト内のアイテムを反復処理し、元のループと同じ方法でプロパティを読み込むが、プロキシ オブジェクトのプロパティを設定する a が含まれるので、最初のループでは段落のテキストをログに記録できません。 `text` `context.sync` `text` `paragraph`</span><span class="sxs-lookup"><span data-stu-id="177ab-147">The first loop iterates through the items in the collection object and loads the `text` property just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object.</span></span> <span data-ttu-id="177ab-148">代わりに、オブジェクトを `paragraph` 配列に追加します。</span><span class="sxs-lookup"><span data-stu-id="177ab-148">Instead, it adds the `paragraph` object to an array.</span></span>
- <span data-ttu-id="177ab-149">2 番目のループは、最初のループによって作成された配列を反復処理し、各アイテム `text` のログを記録 `paragraph` します。</span><span class="sxs-lookup"><span data-stu-id="177ab-149">The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item.</span></span> <span data-ttu-id="177ab-150">これは、2 つのループ `context.sync` の間に含まれるすべてのプロパティが設定されたためです `text` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-150">This is possible because the `context.sync` that came between the two loops populated all the `text` properties.</span></span>

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

<span data-ttu-id="177ab-151">前の例では、a を含むループを分割ループ パターンに変換する手順 `context.sync` を示します。</span><span class="sxs-lookup"><span data-stu-id="177ab-151">The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern.</span></span>

1. <span data-ttu-id="177ab-152">ループを 2 つのループに置き換える。</span><span class="sxs-lookup"><span data-stu-id="177ab-152">Replace the loop with two loops.</span></span>
2. <span data-ttu-id="177ab-153">コレクションを反復処理し、各アイテムを配列に追加し、コードで読み取る必要があるアイテムのプロパティも読み込む最初のループを作成します。</span><span class="sxs-lookup"><span data-stu-id="177ab-153">Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.</span></span>
3. <span data-ttu-id="177ab-154">最初のループに続き、プロキシ `context.sync` オブジェクトに読み込まれたプロパティを設定する呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="177ab-154">Following the first loop, call `context.sync` to populate the proxy objects with any loaded properties.</span></span>
4. <span data-ttu-id="177ab-155">2 番目のループに従って、最初のループで作成された配列を反復処理し、読み込まれた `context.sync` プロパティを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="177ab-155">Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.</span></span>

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a><span data-ttu-id="177ab-156">関連付けオブジェクト パターンを使用してドキュメント内のオブジェクトを処理する</span><span class="sxs-lookup"><span data-stu-id="177ab-156">Processing objects in the document with the correlated objects pattern</span></span>

<span data-ttu-id="177ab-157">コレクション内のアイテムを処理するには、アイテム自体に含されていないデータが必要な、より複雑なシナリオについて考えます。</span><span class="sxs-lookup"><span data-stu-id="177ab-157">Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves.</span></span> <span data-ttu-id="177ab-158">このシナリオでは、テンプレートから作成されたいくつかの定型文を含むドキュメントを操作する Word アドインを想定しています。</span><span class="sxs-lookup"><span data-stu-id="177ab-158">The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text.</span></span> <span data-ttu-id="177ab-159">テキストに散在するインスタンスは、"{Coordinator}"、"{Deputy}"、および "{Manager}" というプレースホルダー文字列の 1 つ以上のインスタンスです。</span><span class="sxs-lookup"><span data-stu-id="177ab-159">Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}".</span></span> <span data-ttu-id="177ab-160">アドインは、各プレースホルダーを一部のユーザーの名前に置き換える。</span><span class="sxs-lookup"><span data-stu-id="177ab-160">The add-in replaces each placeholder with some person's name.</span></span> <span data-ttu-id="177ab-161">この記事では、アドインの UI は重要ではありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-161">The UI of the add-in is not important to this article.</span></span> <span data-ttu-id="177ab-162">たとえば、作業ウィンドウに 3 つのテキスト ボックスが表示され、それぞれにプレースホルダーの 1 つが付きます。</span><span class="sxs-lookup"><span data-stu-id="177ab-162">For example, it could have a task pane with three text boxes, each labeled with one of the placeholders.</span></span> <span data-ttu-id="177ab-163">ユーザーは、各テキスト ボックスに名前を入力し、[置換] ボタンを **押** します。</span><span class="sxs-lookup"><span data-stu-id="177ab-163">The user enters a name in each text box and then presses a **Replace** button.</span></span> <span data-ttu-id="177ab-164">ボタンのハンドラーは、名前をプレースホルダーにマップする配列を作成し、各プレースホルダーを割り当てられた名前に置き換える。</span><span class="sxs-lookup"><span data-stu-id="177ab-164">The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.</span></span> 

<span data-ttu-id="177ab-165">この UI を使用して実際にアドインを作成して、コードを試す必要はありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-165">You don't need to actually produce an add-in with this UI to experiment with the code.</span></span> <span data-ttu-id="177ab-166">このツールを使用[Script Labコード](../overview/explore-with-script-lab.md)のプロトタイプを作成できます。</span><span class="sxs-lookup"><span data-stu-id="177ab-166">You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to prototype the important code.</span></span> <span data-ttu-id="177ab-167">マッピング配列を作成するには、次の代入ステートメントを使用します。</span><span class="sxs-lookup"><span data-stu-id="177ab-167">Use the following assignment statement to create the mapping array.</span></span>

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

<span data-ttu-id="177ab-168">次のコードは、ループ内で使用した場合に、各プレースホルダーを割り当てられた名前に置き換える方法 `context.sync` を示しています。</span><span class="sxs-lookup"><span data-stu-id="177ab-168">The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops.</span></span>

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

<span data-ttu-id="177ab-169">前のコードでは、外部ループと内部ループがあります。</span><span class="sxs-lookup"><span data-stu-id="177ab-169">In the preceding code, there is an outer and an inner loop.</span></span> <span data-ttu-id="177ab-170">これらの各ファイルには、 `context.sync` が含まれる。</span><span class="sxs-lookup"><span data-stu-id="177ab-170">Each of them contains a `context.sync`.</span></span> <span data-ttu-id="177ab-171">この記事の最初のコード スニペットに基づいて、内部ループ内を内部ループの後に移動できる `context.sync` 可能性があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-171">Based on the very first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop.</span></span> <span data-ttu-id="177ab-172">しかし、それでもコードは外側のループに (実際には 2 `context.sync` つ) 残ります。</span><span class="sxs-lookup"><span data-stu-id="177ab-172">But that would still leave the code with a `context.sync` (two of them actually) in the outer loop.</span></span> <span data-ttu-id="177ab-173">次のコードは、ループから削除する `context.sync` 方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="177ab-173">The following code shows how you can remove `context.sync` from the loops.</span></span> <span data-ttu-id="177ab-174">以下のコードについて説明します。</span><span class="sxs-lookup"><span data-stu-id="177ab-174">We discuss the code below.</span></span>

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

<span data-ttu-id="177ab-175">コードでは、分割ループ パターンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="177ab-175">Note the code uses the split loop pattern:</span></span>

- <span data-ttu-id="177ab-176">前の例の外側のループは 2 つに分割されています。</span><span class="sxs-lookup"><span data-stu-id="177ab-176">The outer loop from the preceding example has been split into two.</span></span> <span data-ttu-id="177ab-177">(2 番目のループには内部ループがあります。これは、コードが一連のジョブ (またはプレースホルダー) を反復処理し、そのセット内で一致する範囲を反復処理している場合に予期されます)。</span><span class="sxs-lookup"><span data-stu-id="177ab-177">(The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)</span></span>
- <span data-ttu-id="177ab-178">各メジャー ループ `context.sync` の後に発生しますが、ループ `context.sync` 内には含めはありません。</span><span class="sxs-lookup"><span data-stu-id="177ab-178">There is a `context.sync` after each major loop, but no `context.sync` inside any loop.</span></span>
- <span data-ttu-id="177ab-179">2 番目のメジャー ループは、最初のループで作成された配列を反復処理します。</span><span class="sxs-lookup"><span data-stu-id="177ab-179">The second major loop iterates through an array that is created in the first loop.</span></span>

<span data-ttu-id="177ab-180">ただし、最初のループで作成された配列には、セクション「分割ループ パターンを使用してドキュメントから値を読み取る」セクションで行った最初のループと同様に、Office オブジェクト[だけが含まれます](#reading-values-from-the-document-with-the-split-loop-pattern)。</span><span class="sxs-lookup"><span data-stu-id="177ab-180">But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern).</span></span> <span data-ttu-id="177ab-181">これは、Word Range オブジェクトの処理に必要な情報の一部が Range オブジェクト自体ではなく、配列に含まれるため `jobMapping` です。</span><span class="sxs-lookup"><span data-stu-id="177ab-181">This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.</span></span>

<span data-ttu-id="177ab-182">したがって、最初のループで作成された配列内のオブジェクトは、2 つのプロパティを持つカスタム オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="177ab-182">So, the objects in the array created in the first loop are custom objects that have two properties.</span></span> <span data-ttu-id="177ab-183">1 つ目は、特定の役職 (プレースホルダー文字列) に一致する Word の範囲の配列で、2 つ目はジョブに割り当てられたユーザーの名前を示す文字列です。</span><span class="sxs-lookup"><span data-stu-id="177ab-183">The first is an array of Word Ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job.</span></span> <span data-ttu-id="177ab-184">これにより、指定した範囲を処理するために必要なすべての情報が、その範囲を含む同じカスタム オブジェクトに含まれているため、最終的なループを簡単に記述し、読みやすくします。</span><span class="sxs-lookup"><span data-stu-id="177ab-184">This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range.</span></span> <span data-ttu-id="177ab-185">_**correlatedObject**.rangesMatchingJob.items[j]_ を置き換える必要がある名前は、同じオブジェクトのもう 1 つのプロパティです _**。correlatedObject**.personAssignedToJob_ です。</span><span class="sxs-lookup"><span data-stu-id="177ab-185">The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.</span></span>

<span data-ttu-id="177ab-186">この分割ループ パターンのバリエーションを、相関オブジェクト **パターンと呼** ぶ。</span><span class="sxs-lookup"><span data-stu-id="177ab-186">We call this variation of the split loop pattern the **correlated objects** pattern.</span></span> <span data-ttu-id="177ab-187">一般的な考え方は、最初のループがカスタム オブジェクトの配列を作成する方法です。</span><span class="sxs-lookup"><span data-stu-id="177ab-187">The general idea is that the first loop creates an array of custom objects.</span></span> <span data-ttu-id="177ab-188">各オブジェクトには、コレクション オブジェクト内の項目の 1 つ (またはOfficeの配列) の値を持つプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="177ab-188">Each object has a property whose value is one of the items in an Office collection object (or an array of such items).</span></span> <span data-ttu-id="177ab-189">カスタム オブジェクトには他のプロパティが含まれています。各プロパティには、最終ループ内のオブジェクトを処理するためにOffice情報が提供されます。</span><span class="sxs-lookup"><span data-stu-id="177ab-189">The custom object has other properties, each of which provides information needed to process the Office objects in the final loop.</span></span> <span data-ttu-id="177ab-190">カスタム関連付 [けオブジェクト](#other-examples-of-these-patterns) に複数のプロパティがある例へのリンクについては、「これらのパターンのその他の例」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="177ab-190">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.</span></span>

<span data-ttu-id="177ab-191">もう 1 つの注意点: カスタム相関オブジェクトの配列を作成するために複数のループが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-191">One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects.</span></span> <span data-ttu-id="177ab-192">これは、別のコレクション オブジェクトの処理に使用される情報を収集するために、Office コレクション オブジェクトの各メンバーのプロパティを読み取る必要がある場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="177ab-192">This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object.</span></span> <span data-ttu-id="177ab-193">(たとえば、アドインは、その列のタイトルに基づいていくつかの列のセルに数値形式を適用するつもりなので、Excel テーブル内のすべての列のタイトルを読み取る必要があります)。ただし、ループではなく、ループの間に s を常 `context.sync` に保持できます。</span><span class="sxs-lookup"><span data-stu-id="177ab-193">(For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop.</span></span> <span data-ttu-id="177ab-194">例については、 [セクション「これらのパターンのその他の例](#other-examples-of-these-patterns) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="177ab-194">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.</span></span>

## <a name="other-examples-of-these-patterns"></a><span data-ttu-id="177ab-195">これらのパターンの他の例</span><span class="sxs-lookup"><span data-stu-id="177ab-195">Other examples of these patterns</span></span>

- <span data-ttu-id="177ab-196">ループを使用する Excel の非常に簡単な例については、このスタック オーバーフローの質問に対する受け入れ可能な回答を参照してください。context.sync の前に複数の `Array.forEach` [context.load](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)をキューに入れられますか?</span><span class="sxs-lookup"><span data-stu-id="177ab-196">For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span></span>
- <span data-ttu-id="177ab-197">ループを使用し、構文を使用しない Word の簡単な例については、このスタック オーバーフローの質問に対する受け入れ可能な回答を参照してください `Array.forEach` `async` / `await` [。Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)を使用してコンテンツ コントロールを使用してすべての段落を反復処理します。</span><span class="sxs-lookup"><span data-stu-id="177ab-197">For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span></span>
- <span data-ttu-id="177ab-198">TypeScript で記述されている Word の例については、サンプルの Word アドイン [Angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)スタイル チェッカー (特に [ ument.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)のファイルword.docを参照してください。</span><span class="sxs-lookup"><span data-stu-id="177ab-198">For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span></span> <span data-ttu-id="177ab-199">これは、ループの混合 `for` を `Array.forEach` 持っています。</span><span class="sxs-lookup"><span data-stu-id="177ab-199">It has a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="177ab-200">高度な Word サンプルの場合は、この[gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)を新しいツール[にScript Labします](../overview/explore-with-script-lab.md)。</span><span class="sxs-lookup"><span data-stu-id="177ab-200">For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md).</span></span> <span data-ttu-id="177ab-201">gist を使用するコンテキストについては、テキストの置換後に同期されていないスタック オーバーフローの質問ドキュメントに対する受け入れ可能 [な回答を参照してください](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)。</span><span class="sxs-lookup"><span data-stu-id="177ab-201">For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text).</span></span> <span data-ttu-id="177ab-202">このサンプルでは、3 つのプロパティを持つカスタムの関連付けオブジェクトの種類を作成します。</span><span class="sxs-lookup"><span data-stu-id="177ab-202">This sample creates a custom correlating object type that has three properties.</span></span> <span data-ttu-id="177ab-203">合計 3 つのループを使用して相関オブジェクトの配列を作成し、さらに 2 つのループを使用して最終的な処理を実行します。</span><span class="sxs-lookup"><span data-stu-id="177ab-203">It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing.</span></span> <span data-ttu-id="177ab-204">ループとループの混合 `for` `Array.forEach` があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-204">There are a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="177ab-205">分割ループや相関オブジェクト パターンの厳密な例ではありませんが、セル値のセットを単一の通貨で他の通貨に変換する方法を示す高度な Excel サンプルがあります `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-205">Although not strictly an example of the split loop or correlated objects patterns, there is an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`.</span></span> <span data-ttu-id="177ab-206">このツールを試す場合は、Script Lab [ツールを](../overview/explore-with-script-lab.md)開き **、Currency Converter サンプルに移動** します。</span><span class="sxs-lookup"><span data-stu-id="177ab-206">To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) and navigate to the **Currency Converter** sample.</span></span>

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a><span data-ttu-id="177ab-207">いつこの *記事の* パターンを使う必要がありますか?</span><span class="sxs-lookup"><span data-stu-id="177ab-207">When should you *not* use the patterns in this article?</span></span>

<span data-ttu-id="177ab-208">Excel呼び出しでは、5 MB を超えるデータを読み取る必要があります `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-208">Excel cannot read more than 5 MB of data in a given call of `context.sync`.</span></span> <span data-ttu-id="177ab-209">この制限を超えると、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="177ab-209">If this limit is exceeded, an error is thrown.</span></span> <span data-ttu-id="177ab-210">(詳細については、「Excel アドインのリソース制限とパフォーマンスの最適化」[](resource-limits-and-performance-optimization.md#excel-add-ins)の「Office アドイン」を参照してください。この制限に近づくことは非常にまれですが、アドインでこれが発生する可能性がある場合は、コードですべてのデータを 1 つのループで読み込み、ループに従う必要はありません `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="177ab-210">(See the "Excel add-ins section" of [Resource limits and performance optimization for Office Add-ins](resource-limits-and-performance-optimization.md#excel-add-ins) for more information.) It's very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`.</span></span> <span data-ttu-id="177ab-211">ただし、コレクション オブジェクトに対するループの繰り返しごとに a `context.sync` を使用しないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="177ab-211">But you still should avoid having a `context.sync` in every iteration of a loop over a collection object.</span></span> <span data-ttu-id="177ab-212">代わりに、コレクション内のアイテムのサブセットを定義し、ループ間で各サブセットを順番 `context.sync` にループします。</span><span class="sxs-lookup"><span data-stu-id="177ab-212">Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops.</span></span> <span data-ttu-id="177ab-213">これは、サブセットを反復処理し、これらの外側の各反復に含まれる外部ループ `context.sync` で構成できます。</span><span class="sxs-lookup"><span data-stu-id="177ab-213">You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.</span></span>
