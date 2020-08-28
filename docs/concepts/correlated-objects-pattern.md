---
title: ループで context.sync メソッドを使用しないでください
description: 分割ループと相関オブジェクトのパターンを使用して、コンテキストの呼び出しを回避する方法について説明します。ループでの同期。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: d3628400ef783035cf6a816144dbd5cfb30582ee
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292996"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a><span data-ttu-id="b9235-103">ループで context.sync メソッドを使用しないでください</span><span class="sxs-lookup"><span data-stu-id="b9235-103">Avoid using the context.sync method in loops</span></span>

> [!NOTE]
> <span data-ttu-id="b9235-104">この記事では、 &mdash; バッチシステムを使用して office ドキュメントを操作する、Excel、Word、OneNote、Visio 用の4つのアプリケーション固有の Office JavaScript api のうち、少なくとも1つを操作する最初の段階にとどまらないことを前提としてい &mdash; ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-104">This article assumes that you're beyond the beginning stage of working with at least one of the four application-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document.</span></span> <span data-ttu-id="b9235-105">特に、呼び出しとは何か、コレクションオブジェクトについて理解しておく必要があることを理解しておく必要があり `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-105">In particular, you should know what a call of `context.sync` does and you should know what a collection object is.</span></span> <span data-ttu-id="b9235-106">その段階にない場合は、 [Office JAVASCRIPT API](../develop/understanding-the-javascript-api-for-office.md) と、その記事の「アプリケーション固有」の下にリンクされているドキュメントを理解してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-106">If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "application-specific" in that article.</span></span>

<span data-ttu-id="b9235-107">アプリケーション固有の API モデル (Excel、Word、OneNote、Visio) の1つを使用する Office アドインのプログラミングシナリオによっては、コードでコレクションオブジェクトのすべてのメンバーからいくつかのプロパティを読み取り、書き込み、または処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-107">For some programming scenarios in Office Add-ins that use one of the application-specific API models (for Excel, Word, OneNote, and Visio), your code needs to read, write, or process some property from every member of a collection object.</span></span> <span data-ttu-id="b9235-108">たとえば、特定のテーブル列のすべてのセルの値を取得する必要がある Excel アドイン、またはドキュメント内の文字列のすべてのインスタンスを強調表示する必要がある Word アドイン。</span><span class="sxs-lookup"><span data-stu-id="b9235-108">For example, an Excel add-in that needs to get the values of every cell in a particular table column or a Word add-in that needs to highlight every instance of a string in the document.</span></span> <span data-ttu-id="b9235-109">コレクションオブジェクトのプロパティのメンバーを反復処理する必要があります `items` が、パフォーマンス上の理由から、 `context.sync` ループのすべての反復処理での呼び出しを回避する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-109">You need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you need to avoid calling `context.sync` in every iteration of the loop.</span></span> <span data-ttu-id="b9235-110">のすべての呼び出しで `context.sync` は、アドインから Office ドキュメントへのラウンドトリップがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-110">Every call of `context.sync` is a round trip from the add-in to the Office document.</span></span> <span data-ttu-id="b9235-111">ラウンドトリップがインターネット経由で行われるため、特に web 上の Office でアドインが実行されている場合、ラウンドトリップが繰り返されるとパフォーマンスが低下します。</span><span class="sxs-lookup"><span data-stu-id="b9235-111">Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.</span></span>

> [!NOTE]
> <span data-ttu-id="b9235-112">この記事のすべての例ではループを使用して `for` いますが、ここで説明する操作は、次のような配列を反復処理できる任意のループステートメントに適用されます。</span><span class="sxs-lookup"><span data-stu-id="b9235-112">All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:</span></span>
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> <span data-ttu-id="b9235-113">これらは、関数が渡され、配列内の項目に適用される、次のような配列メソッドにも適用されます。</span><span class="sxs-lookup"><span data-stu-id="b9235-113">They also apply to any array method to which a function is passed and applied to the items in the array, including the following:</span></span>
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

## <a name="writing-to-the-document"></a><span data-ttu-id="b9235-114">ドキュメントへの書き込み</span><span class="sxs-lookup"><span data-stu-id="b9235-114">Writing to the document</span></span>

<span data-ttu-id="b9235-115">最も単純なケースでは、コレクションオブジェクトのメンバーだけが書き込み、プロパティの読み取りは行われません。</span><span class="sxs-lookup"><span data-stu-id="b9235-115">In the simplest case, you are only writing to members of a collection object, not reading their properties.</span></span> <span data-ttu-id="b9235-116">たとえば、次のコードでは、Word 文書内の "the" のすべてのインスタンスが黄色で強調表示されています。</span><span class="sxs-lookup"><span data-stu-id="b9235-116">For example, the following code highlights in yellow every instance of "the" in a Word document.</span></span>

> [!NOTE]
> <span data-ttu-id="b9235-117">通常は、 `context.sync` アプリケーションメソッドの閉じる側の "}" 文字 (、など) の直前に配置することをお勧めし `run` `Excel.run` `Word.run` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-117">It is generally a good practice to put have a final `context.sync` just before the closing "}" character of the application `run` method (such as `Excel.run`, `Word.run`, etc.).</span></span> <span data-ttu-id="b9235-118">これは、メソッドでは、 `run` `context.sync` まだ同期されていないキューに入れられたコマンドがある場合に限り、最後に実行したときと同じ方法で非表示の呼び出しを行うためです。</span><span class="sxs-lookup"><span data-stu-id="b9235-118">This is because the `run` method makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that have not yet been synchronized.</span></span> <span data-ttu-id="b9235-119">この呼び出しが非表示になっていることがわかりやすいため、通常は明示的なを追加することをお勧めし `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-119">The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`.</span></span> <span data-ttu-id="b9235-120">ただし、この記事では、呼び出しを最小限にすることについて説明してい `context.sync` ますが、実際には不要な最終処理を追加する方が混乱してい `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-120">However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`.</span></span> <span data-ttu-id="b9235-121">そのため、この記事では、の最後に同期されていないコマンドがない場合は、このままに `run` します。</span><span class="sxs-lookup"><span data-stu-id="b9235-121">So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.</span></span>

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

<span data-ttu-id="b9235-122">前のコードでは、Word の "the" という単語に200のインスタンスが含まれているドキュメントで完了するために1秒で完了していました。</span><span class="sxs-lookup"><span data-stu-id="b9235-122">The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows.</span></span> <span data-ttu-id="b9235-123">しかし、 `await context.sync();` ループの内側の行がコメントアウトされているときに、ループがコメントアウトされた直後の行であれば、処理には1秒あたり10秒だけかかります。</span><span class="sxs-lookup"><span data-stu-id="b9235-123">But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second.</span></span> <span data-ttu-id="b9235-124">Web 上の Word (ブラウザーとしてのエッジを含む) では、ループ内で同期が行われ、ループの後に同期が5倍高速になるまでに3秒で完了しました。</span><span class="sxs-lookup"><span data-stu-id="b9235-124">In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster.</span></span> <span data-ttu-id="b9235-125">"The" の2000インスタンスが含まれるドキュメントでは、(web 上の Word) 80 秒で、ループの内部で同期が行われ、ループ後の同期では約20倍高速になりました。</span><span class="sxs-lookup"><span data-stu-id="b9235-125">In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.</span></span>

> [!NOTE]
> <span data-ttu-id="b9235-126">同期が同時に実行された場合に、ループ内での同期のバージョンが高速に実行されるかどうかを確認する必要が `await` あります。これは、の前にあるキーワードを削除するだけで行うことができ `context.sync()` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-126">It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`.</span></span> <span data-ttu-id="b9235-127">これにより、ランタイムが同期を開始し、同期が完了するのを待たずに、ループの次の反復処理を直ちに開始することができます。</span><span class="sxs-lookup"><span data-stu-id="b9235-127">This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete.</span></span> <span data-ttu-id="b9235-128">ただし、次の理由から、ループを完全に移動するのではなく、この方法を使用しても問題ありません `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-128">However, this is not as good a solution as moving the `context.sync` out of the loop entirely for these reasons:</span></span>
>
> - <span data-ttu-id="b9235-129">同期バッチジョブのコマンドがキューに入れられているのと同様に、バッチジョブ自体は Office でキューに入れられますが、Office はキュー内に50個を超えるバッチジョブをサポートしません。</span><span class="sxs-lookup"><span data-stu-id="b9235-129">Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue.</span></span> <span data-ttu-id="b9235-130">これ以上トリガーエラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="b9235-130">Any more triggers errors.</span></span> <span data-ttu-id="b9235-131">そのため、ループに50を超える反復がある場合は、キューのサイズを超えている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-131">So, if there are more than 50 iterations in a loop, there is a chance that the queue size is exceeded.</span></span> <span data-ttu-id="b9235-132">反復回数が多いほど、発生する可能性が高くなります。</span><span class="sxs-lookup"><span data-stu-id="b9235-132">The greater the number of iterations, the greater the chance of this happening.</span></span> 
> - <span data-ttu-id="b9235-133">"同時" は同時に意味がありません。</span><span class="sxs-lookup"><span data-stu-id="b9235-133">"Concurrently" does not mean simultaneously.</span></span> <span data-ttu-id="b9235-134">その場合でも、1つの同期操作を実行するよりも、複数の同期操作を実行するのにかかる時間が長くなります。</span><span class="sxs-lookup"><span data-stu-id="b9235-134">It would still take longer to execute multiple synchronization operations than to execute one.</span></span>
> - <span data-ttu-id="b9235-135">同時操作は、開始したときと同じ順序で完了することは保証されません。</span><span class="sxs-lookup"><span data-stu-id="b9235-135">Concurrent operations are not guaranteed to complete in the same order in which they started.</span></span> <span data-ttu-id="b9235-136">前の例では、単語 "the" が強調表示されている順序は重要ではありませんが、コレクション内のアイテムを順番に処理することが重要なシナリオがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-136">In the preceding example, it doesn't matter what order the  word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.</span></span>

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a><span data-ttu-id="b9235-137">分割ループパターンを使用してドキュメントから値を読み取る</span><span class="sxs-lookup"><span data-stu-id="b9235-137">Reading values from the document with the split loop pattern</span></span>

<span data-ttu-id="b9235-138">`context.sync`ループ内のを回避することは、コードがそれぞれを処理するときにコレクションアイテムのプロパティを*読み取る*必要がある場合に、より困難になります。</span><span class="sxs-lookup"><span data-stu-id="b9235-138">Avoiding `context.sync`s inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one.</span></span> <span data-ttu-id="b9235-139">コードで、Word 文書内のすべてのコンテンツコントロールを反復処理し、各コントロールに関連付けられている最初の段落のテキストをログに記録する必要があるとします。</span><span class="sxs-lookup"><span data-stu-id="b9235-139">Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control.</span></span> <span data-ttu-id="b9235-140">プログラミングの instincts によって、コントロールをループ処理し、 `text` 各 (最初の) 段落のプロパティを読み込んで、プロキシの paragraph オブジェクトにドキュメントのテキストを設定する呼び出しを行い、それを記録することができ `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-140">Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it.</span></span> <span data-ttu-id="b9235-141">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="b9235-141">The following is an example.</span></span>

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

<span data-ttu-id="b9235-142">このシナリオでは、ループにが含まれないようにするために、 `context.sync` **スプリットループ** パターンを呼び出すパターンを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-142">In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern.</span></span> <span data-ttu-id="b9235-143">このパターンの具体的な例を参照してから、正式な説明にしてみましょう。</span><span class="sxs-lookup"><span data-stu-id="b9235-143">Let's see a concrete example of the pattern before we get to a formal description of it.</span></span> <span data-ttu-id="b9235-144">スプリットループパターンを前述のコードスニペットに適用する方法は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b9235-144">Here's how the split loop pattern can be applied to the preceding code snippet.</span></span> <span data-ttu-id="b9235-145">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-145">Note the following about this code:</span></span>

- <span data-ttu-id="b9235-146">これで2つのループがあり、どちらのループにも入っているので、 `context.sync` 一方のループ内にはありません `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-146">There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.</span></span>
- <span data-ttu-id="b9235-147">最初のループでは、コレクションオブジェクトのアイテムを反復処理し、 `text` 元のループと同じようにプロパティを読み込みますが、最初のループには、 `context.sync` `text` プロキシオブジェクトのプロパティを設定するためのが含まれていないため、段落テキストをログに記録することはできません `paragraph` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-147">The first loop iterates through the items in the collection object and loads the `text` property just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object.</span></span> <span data-ttu-id="b9235-148">代わりに、 `paragraph` オブジェクトを配列に追加します。</span><span class="sxs-lookup"><span data-stu-id="b9235-148">Instead, it adds the `paragraph` object to an array.</span></span>
- <span data-ttu-id="b9235-149">2番目のループでは、最初のループによって作成された配列を反復処理し、各アイテムのをログに記録し `text` `paragraph` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-149">The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item.</span></span> <span data-ttu-id="b9235-150">これが可能なのは、 `context.sync` 2 つのループの間にあるがすべてのプロパティを設定したためです `text` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-150">This is possible because the `context.sync` that came between the two loops populated all the `text` properties.</span></span>

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

<span data-ttu-id="b9235-151">前の例では、を含むループをスプリットループパターンに入れるために、次の手順を提案してい `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-151">The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern:</span></span> 

1. <span data-ttu-id="b9235-152">ループを2つのループに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9235-152">Replace the loop with two loops.</span></span>
2. <span data-ttu-id="b9235-153">最初のループを作成してコレクションを反復処理し、各項目を配列に追加します。また、コードで読み込む必要のあるアイテムのプロパティも読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b9235-153">Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.</span></span> 
3. <span data-ttu-id="b9235-154">最初のループに従って、 `context.sync` プロキシオブジェクトに読み込み済みのプロパティを設定するように呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b9235-154">Following the first loop, call `context.sync` to populate the proxy objects with any loaded properties.</span></span> 
4. <span data-ttu-id="b9235-155">第2のループを使用して、 `context.sync` 最初のループで作成された配列に対して反復処理を行い、読み込まれたプロパティを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="b9235-155">Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.</span></span>

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a><span data-ttu-id="b9235-156">[相関オブジェクト] パターンを使用してドキュメント内のオブジェクトを処理する</span><span class="sxs-lookup"><span data-stu-id="b9235-156">Processing objects in the document with the correlated objects pattern</span></span>

<span data-ttu-id="b9235-157">コレクション内のアイテムを処理するには、アイテム自体にはないデータが必要になるという、より複雑なシナリオを考えてみましょう。</span><span class="sxs-lookup"><span data-stu-id="b9235-157">Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves.</span></span> <span data-ttu-id="b9235-158">このシナリオでは、テンプレートから作成されたドキュメントに対して何らかの定型テキストを使用して操作する Word アドインを小売します。</span><span class="sxs-lookup"><span data-stu-id="b9235-158">The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text.</span></span> <span data-ttu-id="b9235-159">テキストに分散されているのは、"{コーディネーター}"、"{Deputy}"、および "{Manager}" の各プレースホルダー文字列の1つ以上のインスタンスです。</span><span class="sxs-lookup"><span data-stu-id="b9235-159">Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}".</span></span> <span data-ttu-id="b9235-160">各プレースホルダーは、アドインによってユーザーの名前に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="b9235-160">The add-in replaces each placeholder with some person's name.</span></span> <span data-ttu-id="b9235-161">この記事では、アドインの UI は重要ではありません。</span><span class="sxs-lookup"><span data-stu-id="b9235-161">The UI of the add-in is not important to this article.</span></span> <span data-ttu-id="b9235-162">たとえば、3つのテキストボックスを含む作業ウィンドウがあり、それぞれにプレースホルダーのいずれかでラベルが付けられているとします。</span><span class="sxs-lookup"><span data-stu-id="b9235-162">For example, it could have a task pane with three text boxes, each labeled with one of the placeholders.</span></span> <span data-ttu-id="b9235-163">ユーザーは、各テキストボックスに名前を入力し、[ **置換** ] ボタンを押します。</span><span class="sxs-lookup"><span data-stu-id="b9235-163">The user enters a name in each text box and then presses a **Replace** button.</span></span> <span data-ttu-id="b9235-164">ボタンのハンドラーは、名前をプレースホルダーにマップする配列を作成し、各プレースホルダーを割り当てられた名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9235-164">The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.</span></span> 

<span data-ttu-id="b9235-165">コードを試すために、この UI を使用してアドインを実際に作成する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="b9235-165">You don't need to actually produce an add-in with this UI to experiment with the code.</span></span> <span data-ttu-id="b9235-166">[スクリプトラボツール](../overview/explore-with-script-lab.md)を使用して、重要なコードのプロトタイプを作成できます。</span><span class="sxs-lookup"><span data-stu-id="b9235-166">You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to prototype the important code.</span></span> <span data-ttu-id="b9235-167">次の代入ステートメントを使用して、マッピング配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="b9235-167">Use the following assignment statement to create the mapping array.</span></span>

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

<span data-ttu-id="b9235-168">次のコードは、内部ループを使用した場合に、各プレースホルダーを割り当てられた名前に置き換える方法を示して `context.sync` います。</span><span class="sxs-lookup"><span data-stu-id="b9235-168">The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops.</span></span>

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

<span data-ttu-id="b9235-169">上記のコードでは、外部ループと内部ループがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-169">In the preceding code, there is an outer and an inner loop.</span></span> <span data-ttu-id="b9235-170">各には、が含まれてい `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-170">Each of them contains a `context.sync`.</span></span> <span data-ttu-id="b9235-171">この記事の最初のコードスニペットに基づいて、内側のループの後に内側のループを単純に移動できることがわかるでしょう `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-171">Based on the very first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop.</span></span> <span data-ttu-id="b9235-172">しかし、その場合でも、外側の `context.sync` ループには (実際にはそのうちの2つの) コードを残しておきます。</span><span class="sxs-lookup"><span data-stu-id="b9235-172">But that would still leave the code with a `context.sync` (two of them actually) in the outer loop.</span></span> <span data-ttu-id="b9235-173">次のコードは、ループから削除する方法を示して `context.sync` います。</span><span class="sxs-lookup"><span data-stu-id="b9235-173">The following code shows how you can remove `context.sync` from the loops.</span></span> <span data-ttu-id="b9235-174">次のコードについて説明します。</span><span class="sxs-lookup"><span data-stu-id="b9235-174">We discuss the code below.</span></span>

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

<span data-ttu-id="b9235-175">メモコードでは、分割ループパターンを使用しています。</span><span class="sxs-lookup"><span data-stu-id="b9235-175">Note the code uses the split loop pattern:</span></span>

- <span data-ttu-id="b9235-176">前の例の外側のループは、2つに分割されています。</span><span class="sxs-lookup"><span data-stu-id="b9235-176">The outer loop from the preceding example has been split into two.</span></span> <span data-ttu-id="b9235-177">(2 番目のループには内側のループがあります。これは、コードが一連のジョブ (またはプレースホルダー) を反復処理しており、そのセット内で一致する範囲を反復処理しているためです)。</span><span class="sxs-lookup"><span data-stu-id="b9235-177">(The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)</span></span>
- <span data-ttu-id="b9235-178">各メジャーループの後にはがあり `context.sync` ますが、ループ内にはありません `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-178">There is a `context.sync` after each major loop, but no `context.sync` inside any loop.</span></span>
- <span data-ttu-id="b9235-179">2番目のメジャーループでは、最初のループで作成された配列を反復処理します。</span><span class="sxs-lookup"><span data-stu-id="b9235-179">The second major loop iterates through an array that is created in the first loop.</span></span>

<span data-ttu-id="b9235-180">ただし、最初のループで作成された配列には、[分割ループパターンを使用してドキュメントから値を読み込む](#reading-values-from-the-document-with-the-split-loop-pattern)セクションの最初のループと同じように、Office オブジェクトのみが含まれているわけでは*ありません*。</span><span class="sxs-lookup"><span data-stu-id="b9235-180">But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern).</span></span> <span data-ttu-id="b9235-181">これは、Word の Range オブジェクトの処理に必要な情報の一部は、Range オブジェクト自体ではなく、配列から取得されるためです `jobMapping` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-181">This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.</span></span>

<span data-ttu-id="b9235-182">そのため、最初のループで作成された配列内のオブジェクトは、2つのプロパティを持つカスタムオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="b9235-182">So, the objects in the array created in the first loop are custom objects that have two properties.</span></span> <span data-ttu-id="b9235-183">1つ目は、特定の役職 (つまり、プレースホルダー文字列) に一致する単語範囲の配列で、2番目の配列は、ジョブに割り当てられた人物の名前を提供する文字列です。</span><span class="sxs-lookup"><span data-stu-id="b9235-183">The first is an array of Word Ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job.</span></span> <span data-ttu-id="b9235-184">これにより、指定された範囲を処理するために必要なすべての情報が範囲を含む同じカスタムオブジェクトに格納されるため、最終ループが簡単に書き込み可能になり、読みやすくなります。</span><span class="sxs-lookup"><span data-stu-id="b9235-184">This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range.</span></span> <span data-ttu-id="b9235-185">CorrelatedObject を置き換える名前。 _ **correlatedObject**rangesMatchingJob [j]_ は、同じオブジェクトのもう1つのプロパティです。 _ **correlatedObject**: personAssignedToJob_。</span><span class="sxs-lookup"><span data-stu-id="b9235-185">The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.</span></span>

<span data-ttu-id="b9235-186">このような分割ループパターンは、このバリエーションによって関連付けられた **オブジェクト** パターンを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b9235-186">We call this variation of the split loop pattern the **correlated objects** pattern.</span></span> <span data-ttu-id="b9235-187">一般的な考え方は、最初のループでカスタムオブジェクトの配列が作成されるということです。</span><span class="sxs-lookup"><span data-stu-id="b9235-187">The general idea is that the first loop creates an array of custom objects.</span></span> <span data-ttu-id="b9235-188">各オブジェクトには、Office コレクションオブジェクト内の項目のいずれか (またはそのような項目の配列) の値を持つプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-188">Each object has a property whose value is one of the items in an Office collection object (or an array of such items).</span></span> <span data-ttu-id="b9235-189">カスタムオブジェクトには、最後のループで Office オブジェクトを処理するために必要な情報が含まれている他のプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-189">The custom object has other properties, each of which provides information needed to process the Office objects in the final loop.</span></span> <span data-ttu-id="b9235-190">カスタム関連付けオブジェクトに3つ以上のプロパティがある場合のリンクについては、「 [その他のパターンの例](#other-examples-of-these-patterns) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-190">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.</span></span>

<span data-ttu-id="b9235-191">さらに注意してください。カスタムの関連付けが必要なオブジェクトの配列を作成するためだけに複数のループが発生する場合があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-191">One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects.</span></span> <span data-ttu-id="b9235-192">これは、ある Office コレクションオブジェクトの各メンバーのプロパティを読み込んで、別のコレクションオブジェクトの処理に使用される情報を収集する必要がある場合に発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="b9235-192">This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object.</span></span> <span data-ttu-id="b9235-193">(たとえば、アドインでは、列のタイトルに基づいて一部の列のセルに数値の書式を適用するため、コードでは、Excel のテーブル内のすべての列のタイトルを読み取る必要があります)。ただし、ループではなく、ループ間では常に s を続けることができ `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-193">(For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop.</span></span> <span data-ttu-id="b9235-194">[このようなパターンの](#other-examples-of-these-patterns)例については、「その他の例」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-194">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.</span></span>

## <a name="other-examples-of-these-patterns"></a><span data-ttu-id="b9235-195">これらのパターンのその他の例</span><span class="sxs-lookup"><span data-stu-id="b9235-195">Other examples of these patterns</span></span>

- <span data-ttu-id="b9235-196">ループを使用する Excel の非常に簡単な例につい `Array.forEach` ては、「このスタックオーバーフローに対する応答の受け入れ」を参照してください。複数の[コンテキストをキューに追加することができます。コンテキストを同期する前に読み込みます](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)。</span><span class="sxs-lookup"><span data-stu-id="b9235-196">For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span></span>
- <span data-ttu-id="b9235-197">ループを使用し、構文を使用しない Word の簡単な例につい `Array.forEach` `async` / `await` ては、「このスタックオーバーフローの回答」を参照してください。 [Office JavaScript API を使用したコンテンツコントロールを含むすべての段落](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)に対して反復処理を行います。</span><span class="sxs-lookup"><span data-stu-id="b9235-197">For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span></span>
- <span data-ttu-id="b9235-198">TypeScript で記述されている Word の例については、サンプル [Word アドインの Angular2 スタイルチェッカー](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)(特に、 [ ument ファイルword.doc](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-198">For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span></span> <span data-ttu-id="b9235-199">このメソッドは、and ループが混在してい `for` `Array.forEach` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-199">It has a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="b9235-200">高度な Word サンプルの場合は、 [この gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) を [スクリプトラボツール](../overview/explore-with-script-lab.md)にインポートします。</span><span class="sxs-lookup"><span data-stu-id="b9235-200">For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md).</span></span> <span data-ttu-id="b9235-201">Gist を使用するコンテキストについては、「 [テキストの置換後に同期されていない](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)スタックオーバーフローの質問」ドキュメントへの受け入れられた応答を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9235-201">For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text).</span></span> <span data-ttu-id="b9235-202">この例では、3つのプロパティを持つ、関連付けができるカスタムオブジェクトの種類を作成します。</span><span class="sxs-lookup"><span data-stu-id="b9235-202">This sample creates a custom correlating object type that has three properties.</span></span> <span data-ttu-id="b9235-203">合計3つのループを使用して、相関オブジェクトの配列を作成し、さらに2つのループを使用して最終的な処理を行います。</span><span class="sxs-lookup"><span data-stu-id="b9235-203">It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing.</span></span> <span data-ttu-id="b9235-204">とループが混在し `for` てい `Array.forEach` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-204">There are a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="b9235-205">スプリットループや相関オブジェクトのパターンの例は厳密ではありませんが、セルの値のセットを1つだけの他の通貨に変換する方法を示す高度な Excel サンプルがあり `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-205">Although not strictly an example of the split loop or correlated objects patterns, there is an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`.</span></span> <span data-ttu-id="b9235-206">これを実行するには、 [スクリプトラボツール](../overview/explore-with-script-lab.md) を開き、 **通貨コンバーター** のサンプルに移動します。</span><span class="sxs-lookup"><span data-stu-id="b9235-206">To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) and navigate to the **Currency Converter** sample.</span></span>

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a><span data-ttu-id="b9235-207">この記事のパターンを使用し *ない* 場合は、どうすればよいですか。</span><span class="sxs-lookup"><span data-stu-id="b9235-207">When should you *not* use the patterns in this article?</span></span>

<span data-ttu-id="b9235-208">Excel は、指定された通話で 5 MB を超えるデータを読み取ることができません `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="b9235-208">Excel cannot read more than 5 MB of data in a given call of `context.sync`.</span></span> <span data-ttu-id="b9235-209">この制限を超えると、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="b9235-209">If this limit is exceeded, an error is thrown.</span></span> <span data-ttu-id="b9235-210">(詳細については、「 [Office アドインのリソースの制限とパフォーマンスの最適化](resource-limits-and-performance-optimization.md#excel-add-ins) 」の「Excel アドイン」セクションを参照してください)。この制限が適用されることは非常にまれですが、これがアドインで発生する可能性がある場合は、コードですべてのデータを1つのループ *にロードし* て、ループに従う必要があり `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-210">(See the "Excel add-ins section" of [Resource limits and performance optimization for Office Add-ins](resource-limits-and-performance-optimization.md#excel-add-ins) for more information.) It's very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`.</span></span> <span data-ttu-id="b9235-211">しかし、 `context.sync` コレクションオブジェクトに対するループの繰り返しが発生しないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9235-211">But you still should avoid having a `context.sync` in every iteration of a loop over a collection object.</span></span> <span data-ttu-id="b9235-212">代わりに、ループ間にを使用して、コレクション内の項目のサブセットを定義し、各サブセットを順番にループし `context.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="b9235-212">Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops.</span></span> <span data-ttu-id="b9235-213">これは、サブセットを反復処理する外側のループを使用して構造化し、 `context.sync` これらの外側の反復のそれぞれにを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="b9235-213">You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.</span></span>
