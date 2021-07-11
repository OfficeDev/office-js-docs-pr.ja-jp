---
title: プレゼンテーションにスライドをPowerPointする
description: プレゼンテーションから別のプレゼンテーションにスライドを挿入する方法について説明します。
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 9b106e8940e7b0f19678e0467d8e900ffecd9438
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348784"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="2d520-103">プレゼンテーションにスライドをPowerPointする</span><span class="sxs-lookup"><span data-stu-id="2d520-103">Insert slides in a PowerPoint presentation</span></span>

<span data-ttu-id="2d520-104">1 PowerPoint PowerPointアドインは、アプリケーション固有の JavaScript ライブラリを使用して、1 つのプレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="2d520-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="2d520-105">挿入されたスライドがソース プレゼンテーションの書式設定を保持するか、ターゲット プレゼンテーションの書式設定を保持するかどうかを制御できます。</span><span class="sxs-lookup"><span data-stu-id="2d520-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span>

<span data-ttu-id="2d520-106">スライド挿入 API は、主にプレゼンテーション テンプレートのシナリオで使用されます。既知のプレゼンテーションは、アドインによって挿入できるスライドのプールとして機能します。</span><span class="sxs-lookup"><span data-stu-id="2d520-106">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="2d520-107">このようなシナリオでは、ユーザーまたは顧客のどちらかが、スライドのタイトルや画像などの選択基準とスライドの ID を関連付けるデータ ソースを作成および管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-107">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="2d520-108">API は、ユーザーが任意のプレゼンテーションからスライドを挿入できるシナリオでも使用できますが、そのシナリオでは、ユーザーは実質的にソース プレゼンテーションからすべてのスライドを挿入する制限があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-108">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="2d520-109">詳細 [については、「挿入するスライドの選択](#selecting-which-slides-to-insert) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d520-109">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="2d520-110">プレゼンテーションから別のプレゼンテーションにスライドを挿入するには、2 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-110">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="2d520-111">ソース プレゼンテーション ファイル (.pptx) を base64 形式の文字列に変換します。</span><span class="sxs-lookup"><span data-stu-id="2d520-111">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="2d520-112">base64 ファイルから現在のプレゼンテーションに 1 つ以上のスライドを挿入するには、この `insertSlidesFromBase64` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="2d520-112">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="2d520-113">ソース プレゼンテーションを base64 に変換する</span><span class="sxs-lookup"><span data-stu-id="2d520-113">Convert the source presentation to base64</span></span>

<span data-ttu-id="2d520-114">ファイルを base64 に変換する方法は多数あります。</span><span class="sxs-lookup"><span data-stu-id="2d520-114">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="2d520-115">使用するプログラミング言語とライブラリ、およびアドインのサーバー側またはクライアント側で変換するかどうかは、シナリオによって決まります。</span><span class="sxs-lookup"><span data-stu-id="2d520-115">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="2d520-116">最も一般的には [、FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側の JavaScript で変換を行います。</span><span class="sxs-lookup"><span data-stu-id="2d520-116">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="2d520-117">次の例は、このプラクティスを示しています。</span><span class="sxs-lookup"><span data-stu-id="2d520-117">The following example shows this practice.</span></span>

1. <span data-ttu-id="2d520-118">まず、ソース ファイルへの参照を取得PowerPointします。</span><span class="sxs-lookup"><span data-stu-id="2d520-118">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="2d520-119">この例では、種類のコントロールを `<input>` 使用 `file` して、ユーザーにファイルの選択を求めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-119">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="2d520-120">アドイン ページに次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2d520-120">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="2d520-121">このマークアップは、次のスクリーンショットの UI をページに追加します。</span><span class="sxs-lookup"><span data-stu-id="2d520-121">This markup adds the UI in the following screenshot to the page.</span></span>

    ![HTML ファイルの種類の入力コントロールの前に「スライドを挿入するプレゼンテーションを選択する」というPowerPointを示すスクリーンショット。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="2d520-124">ファイルを取得する方法は他にPowerPointがあります。</span><span class="sxs-lookup"><span data-stu-id="2d520-124">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="2d520-125">たとえば、ファイルがサーバーまたはサーバーに保存されているOneDrive SharePoint、Microsoft Graphを使用してダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="2d520-125">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="2d520-126">詳細については[、「Microsoft Graph](/graph/api/resources/onedrive)ファイルの操作」および「Access Files with [Microsoft Graph」を参照してください](/learn/modules/msgraph-access-file-data/)。</span><span class="sxs-lookup"><span data-stu-id="2d520-126">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="2d520-127">次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当 `change` てる。</span><span class="sxs-lookup"><span data-stu-id="2d520-127">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="2d520-128">(次の手順 `storeFileAsBase64` で関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="2d520-128">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="2d520-129">次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2d520-129">Add the following code.</span></span> <span data-ttu-id="2d520-130">このコードについては以下の点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="2d520-130">Note the following about this code.</span></span>

    - <span data-ttu-id="2d520-131">この `reader.readAsDataURL` メソッドは、ファイルを base64 に変換し、プロパティに格納 `reader.result` します。</span><span class="sxs-lookup"><span data-stu-id="2d520-131">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="2d520-132">メソッドが完了すると、イベント ハンドラーが `onload` トリガーされます。</span><span class="sxs-lookup"><span data-stu-id="2d520-132">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="2d520-133">イベント `onload` ハンドラーは、エンコードされたファイルのメタデータをトリミングし、エンコードされた文字列をグローバル変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="2d520-133">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="2d520-134">base64 でエンコードされた文字列は、後の手順で作成した別の関数によって読み取りを行うので、グローバルに格納されます。</span><span class="sxs-lookup"><span data-stu-id="2d520-134">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="2d520-135">insertSlidesFromBase64 を使用してスライドを挿入する</span><span class="sxs-lookup"><span data-stu-id="2d520-135">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="2d520-136">アドインは[、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)メソッドを使用してPowerPointプレゼンテーションから現在のプレゼンテーションにスライドを挿入します。</span><span class="sxs-lookup"><span data-stu-id="2d520-136">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="2d520-137">次に示すのは、ソース プレゼンテーションのすべてのスライドが現在のプレゼンテーションの先頭に挿入され、挿入されたスライドがソース ファイルの書式を保持する簡単な例です。</span><span class="sxs-lookup"><span data-stu-id="2d520-137">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="2d520-138">これは、base64 でエンコードされたバージョンのプレゼンテーション ファイルを保持する `chosenFileBase64` PowerPoint注意してください。</span><span class="sxs-lookup"><span data-stu-id="2d520-138">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="2d520-139">[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを 2 番目のパラメーターとして渡して、スライドが挿入される場所、ソースまたはターゲットの書式設定を取得するかどうかなど、挿入結果のいくつかの側面を `insertSlidesFromBase64` 制御できます。</span><span class="sxs-lookup"><span data-stu-id="2d520-139">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="2d520-140">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-140">The following is an example.</span></span> <span data-ttu-id="2d520-141">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2d520-141">About this code, note:</span></span>

- <span data-ttu-id="2d520-142">プロパティには `formatting` 、"UseDestinationTheme" と "KeepSourceFormatting" の 2 つの値があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-142">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="2d520-143">必要に応じて、列挙型 (例: ) を `InsertSlideFormatting` 使用できます `PowerPoint.InsertSlideFormatting.useDestinationTheme` 。</span><span class="sxs-lookup"><span data-stu-id="2d520-143">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="2d520-144">この関数は、プロパティで指定されたスライドの直後に、ソース プレゼンテーションからスライドを挿入 `targetSlideId` します。</span><span class="sxs-lookup"><span data-stu-id="2d520-144">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="2d520-145">このプロパティの値は **、*nnn*#** *#* 、\* mmmmmmm\*\*\*、または \**_nnn_ #* mmm\*\*\*のいずれかの文字列で *、nnn* はスライドの ID (通常は 3 桁) で *、mmmmmmm は* スライドの作成 ID (通常は 9 桁) です。</span><span class="sxs-lookup"><span data-stu-id="2d520-145">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="2d520-146">いくつかの例は `267#763315295` 、、 `267#` 、および `#763315295` です。</span><span class="sxs-lookup"><span data-stu-id="2d520-146">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

<span data-ttu-id="2d520-147">もちろん、通常、ターゲット スライドの ID または作成 ID はコーディング時にはわかりません。</span><span class="sxs-lookup"><span data-stu-id="2d520-147">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="2d520-148">より一般的には、アドインはユーザーにターゲット スライドの選択を求める場合があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-148">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="2d520-149">次の手順では、現在選択されているスライドの \***nnn\*#** ID を取得し、それをターゲット スライドとして使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-149">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="2d520-150">共通 JavaScript API のOffice.context.doc[ ument.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用して、現在選択されているスライドの ID を取得する関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="2d520-150">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="2d520-151">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-151">The following is an example.</span></span> <span data-ttu-id="2d520-152">呼び出しは Promise 戻り関数 `getSelectedDataAsync` に埋め込まれている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2d520-152">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="2d520-153">これを行う理由と方法の詳細については [、「Promise-returning 関数でCommon-APIsラップ」を参照してください](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。</span><span class="sxs-lookup"><span data-stu-id="2d520-153">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. <span data-ttu-id="2d520-154">main 関数の[PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_)内で新しい関数を呼び出し、返す ID ("#" 記号と連結) をパラメーターのプロパティの値として渡 `targetSlideId` `InsertSlideOptions` します。</span><span class="sxs-lookup"><span data-stu-id="2d520-154">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="2d520-155">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-155">The following is an example.</span></span>

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="2d520-156">挿入するスライドの選択</span><span class="sxs-lookup"><span data-stu-id="2d520-156">Selecting which slides to insert</span></span>

<span data-ttu-id="2d520-157">[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)パラメーターを使用して、ソース プレゼンテーションから挿入されるスライドを制御することもできます。</span><span class="sxs-lookup"><span data-stu-id="2d520-157">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="2d520-158">これを行うには、ソース プレゼンテーションのスライドの ID の配列をプロパティに割り当 `sourceSlideIds` てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-158">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="2d520-159">次に、4 つのスライドを挿入する例を示します。</span><span class="sxs-lookup"><span data-stu-id="2d520-159">The following is an example that inserts four slides.</span></span> <span data-ttu-id="2d520-160">配列内の各文字列は、プロパティに使用されるパターンの 1 つ以上に従う必要 `targetSlideId` があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-160">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> <span data-ttu-id="2d520-161">スライドは、配列に表示される順序に関係なく、ソース プレゼンテーションに表示されるのと同じ相対順序で挿入されます。</span><span class="sxs-lookup"><span data-stu-id="2d520-161">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="2d520-162">ユーザーがソース プレゼンテーションでスライドの ID または作成 ID を検出できる実用的な方法はありません。</span><span class="sxs-lookup"><span data-stu-id="2d520-162">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="2d520-163">このため、コーディング時にソースの ID を知っている場合、またはアドインが実行時に一部のデータ ソースから取得できる場合にのみ、このプロパティを `sourceSlideIds` 使用できます。</span><span class="sxs-lookup"><span data-stu-id="2d520-163">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="2d520-164">ユーザーがスライド ID を記憶できないので、ユーザーがスライド (タイトルや画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。</span><span class="sxs-lookup"><span data-stu-id="2d520-164">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="2d520-165">したがって、このプロパティは主にプレゼンテーション テンプレートのシナリオで使用されます。アドインは、挿入できるスライドのプールとして機能する特定のプレゼンテーション セットを操作するように `sourceSlideIds` 設計されています。</span><span class="sxs-lookup"><span data-stu-id="2d520-165">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="2d520-166">このようなシナリオでは、ユーザーまたは顧客のどちらかが、選択基準 (タイトルや画像など) と、可能なソース プレゼンテーションのセットから構築されたスライドの ID またはスライド作成の ID を関連付けるデータ ソースを作成および管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d520-166">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>
