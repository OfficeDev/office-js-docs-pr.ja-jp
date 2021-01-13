---
title: PowerPoint プレゼンテーションのスライドの挿入と削除
description: プレゼンテーション間でスライドを挿入する方法と、スライドを削除する方法について説明します。
ms.date: 01/08/2021
localization_priority: Normal
ms.openlocfilehash: a9a4b2efd1e970d9c45885f9a17046bec4de7e72
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839720"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="66e65-103">PowerPoint プレゼンテーションのスライドの挿入と削除</span><span class="sxs-lookup"><span data-stu-id="66e65-103">Insert and delete slides in a PowerPoint presentation</span></span>

<span data-ttu-id="66e65-104">PowerPoint アドインは、PowerPoint のアプリケーション固有の JavaScript ライブラリを使用して、1 つのプレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="66e65-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="66e65-105">挿入したスライドで、元のプレゼンテーションの書式を維持するか、またはターゲット プレゼンテーションの書式設定を維持するかは制御できます。</span><span class="sxs-lookup"><span data-stu-id="66e65-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="66e65-106">プレゼンテーションからスライドを削除できます。</span><span class="sxs-lookup"><span data-stu-id="66e65-106">You can also delete slides from the presentation.</span></span>

<span data-ttu-id="66e65-107">スライド挿入 API は、主にプレゼンテーション テンプレートのシナリオで使用されます。アドインによって挿入できるスライドのプールとして機能する、いくつかの既知のプレゼンテーションがあります。</span><span class="sxs-lookup"><span data-stu-id="66e65-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="66e65-108">このようなシナリオでは、自分または顧客のどちらかが、選択条件 (スライド タイトルや画像など) とスライドの ID を関連付けるデータ ソースを作成および維持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="66e65-109">この API は、ユーザーが任意のプレゼンテーションからスライドを挿入できるシナリオでも使用できますが、そのシナリオでは、ユーザーは実質的にソース プレゼンテーションからすべてのスライドを挿入する方法に制限されます。</span><span class="sxs-lookup"><span data-stu-id="66e65-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="66e65-110">詳細 [については、「挿入するスライドの選択](#selecting-which-slides-to-insert) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="66e65-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="66e65-111">1 つのプレゼンテーションから別のプレゼンテーションにスライドを挿入するには、2 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="66e65-112">ソース プレゼンテーション ファイル (.pptx) を base64 形式の文字列に変換します。</span><span class="sxs-lookup"><span data-stu-id="66e65-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="66e65-113">Base64 ファイルから現在のプレゼンテーションに 1 つ以上のスライドを挿入するには、この `insertSlidesFromBase64` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="66e65-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="66e65-114">ソース プレゼンテーションを base64 に変換する</span><span class="sxs-lookup"><span data-stu-id="66e65-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="66e65-115">ファイルを base64 に変換する方法は多数あります。</span><span class="sxs-lookup"><span data-stu-id="66e65-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="66e65-116">使用するプログラミング言語とライブラリ、およびアドインのサーバー側とクライアント側のどちらで変換するかは、シナリオによって決まります。</span><span class="sxs-lookup"><span data-stu-id="66e65-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="66e65-117">ほとんどの場合 [、FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側で JavaScript で変換を行います。</span><span class="sxs-lookup"><span data-stu-id="66e65-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="66e65-118">次の例は、この方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="66e65-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="66e65-119">まず、ソース PowerPoint ファイルへの参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="66e65-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="66e65-120">この例では、種類のコントロールを `<input>` 使用して、ユーザーにファイルの選択を求 `file` めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="66e65-121">次のマークアップをアドイン ページに追加します。</span><span class="sxs-lookup"><span data-stu-id="66e65-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="66e65-122">このマークアップは、次のスクリーンショットの UI をページに追加します。</span><span class="sxs-lookup"><span data-stu-id="66e65-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![HTML ファイルの種類の入力コントロールの前に、「スライドを挿入する PowerPoint プレゼンテーションを選択する」という説明文が表示されているスクリーンショット](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="66e65-125">PowerPoint ファイルを取得する方法は他に多数あります。</span><span class="sxs-lookup"><span data-stu-id="66e65-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="66e65-126">たとえば、ファイルが OneDrive または SharePoint に保存されている場合は、Microsoft Graph を使用してダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="66e65-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="66e65-127">詳細については [、「Microsoft Graph でのファイルの操作」および「Microsoft Graph](/graph/api/resources/onedrive) を使用した [ファイルへのアクセス」を参照してください](/learn/modules/msgraph-access-file-data/)。</span><span class="sxs-lookup"><span data-stu-id="66e65-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="66e65-128">次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当 `change` てる。</span><span class="sxs-lookup"><span data-stu-id="66e65-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="66e65-129">(次の手順 `storeFileAsBase64` で関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="66e65-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="66e65-130">次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="66e65-130">Add the following code.</span></span> <span data-ttu-id="66e65-131">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="66e65-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="66e65-132">この `reader.readAsDataURL` メソッドは、ファイルを base64 に変換し、プロパティに格納 `reader.result` します。</span><span class="sxs-lookup"><span data-stu-id="66e65-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="66e65-133">メソッドが完了すると、イベント ハンドラーが `onload` トリガーされます。</span><span class="sxs-lookup"><span data-stu-id="66e65-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="66e65-134">イベント ハンドラーは、エンコードされたファイルからメタデータをトリミングし、エンコードされた文字列をグローバル `onload` 変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="66e65-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="66e65-135">base64 でエンコードされた文字列は、後の手順で作成する別の関数によって読み取りが行なうので、グローバルに格納されます。</span><span class="sxs-lookup"><span data-stu-id="66e65-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="66e65-136">insertSlidesFromBase64 を使用してスライドを挿入する</span><span class="sxs-lookup"><span data-stu-id="66e65-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="66e65-137">アドインは [、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) メソッドを使用して、別の PowerPoint プレゼンテーションのスライドを現在のプレゼンテーションに挿入します。</span><span class="sxs-lookup"><span data-stu-id="66e65-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="66e65-138">次に、ソース プレゼンテーションのすべてのスライドを現在のプレゼンテーションの先頭に挿入し、挿入したスライドにソース ファイルの書式を保持する簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="66e65-139">これは、PowerPoint プレゼンテーション ファイルの base64 エンコード バージョンを保持するグローバル変数 `chosenFileBase64` です。</span><span class="sxs-lookup"><span data-stu-id="66e65-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="66e65-140">[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを 2 番目のパラメーターとして渡して、挿入結果の一部の側面 (スライドを挿入する場所、ソースまたはターゲットの書式設定を取得するかどうかを含む) を制御できます `insertSlidesFromBase64` 。</span><span class="sxs-lookup"><span data-stu-id="66e65-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="66e65-141">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-141">The following is an example.</span></span> <span data-ttu-id="66e65-142">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="66e65-142">About this code, note:</span></span>

- <span data-ttu-id="66e65-143">プロパティには `formatting` 、"UseDestinationTheme" と "KeepSourceFormatting" の 2 つの値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="66e65-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="66e65-144">必要に応じて、列挙型 `InsertSlideFormatting` (例: ) を使用できます `PowerPoint.InsertSlideFormatting.useDestinationTheme` 。</span><span class="sxs-lookup"><span data-stu-id="66e65-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="66e65-145">この関数は、プロパティで指定されたスライドの直後に、ソース プレゼンテーションのスライドを挿入 `targetSlideId` します。</span><span class="sxs-lookup"><span data-stu-id="66e65-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="66e65-146">このプロパティの値は **、*nnn*#**、\* *#* mmm\*\*\*、または \**_nnn_ #* mmm\*\*\*の 3 つの形式のいずれかの文字列です。ここで *、nnn* はスライドの ID (通常は 3 桁) で *、mmmmm は* スライドの作成 ID (通常は 9 桁) です。</span><span class="sxs-lookup"><span data-stu-id="66e65-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="66e65-147">たとえば、, `267#763315295` `267#` , and `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="66e65-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="66e65-148">もちろん、通常、ターゲット スライドの ID または作成 ID はコーディング時にはわかりません。</span><span class="sxs-lookup"><span data-stu-id="66e65-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="66e65-149">多くの場合、アドインはユーザーにターゲット スライドの選択を求める場合があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="66e65-150">次の手順は、現在選択されているスライドの \***nnn\*#** ID を取得し、ターゲット スライドとして使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="66e65-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="66e65-151">共通 JavaScript API の [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用して、現在選択されているスライドの ID を取得する関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="66e65-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="66e65-152">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-152">The following is an example.</span></span> <span data-ttu-id="66e65-153">呼び出しは `getSelectedDataAsync` Promise を返す関数に埋め込まれている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="66e65-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="66e65-154">これを行う理由と方法の詳細については、「Promise を返す関数で Common-APIs [をラップする」を参照してください](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。</span><span class="sxs-lookup"><span data-stu-id="66e65-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="66e65-155">メイン関数の[PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_)内で新しい関数を呼び出し、パラメーターのプロパティの値として返される ID ("#" 記号と連結) を渡します。 `targetSlideId` `InsertSlideOptions`</span><span class="sxs-lookup"><span data-stu-id="66e65-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="66e65-156">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="66e65-157">挿入するスライドの選択</span><span class="sxs-lookup"><span data-stu-id="66e65-157">Selecting which slides to insert</span></span>

<span data-ttu-id="66e65-158">[InsertSlideOptions パラメーター](/javascript/api/powerpoint/powerpoint.insertslideoptions)を使用して、ソース プレゼンテーションから挿入するスライドを制御することもできます。</span><span class="sxs-lookup"><span data-stu-id="66e65-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="66e65-159">これを行うには、元のプレゼンテーションのスライドのスライドの配列をプロパティに割り当 `sourceSlideIds` てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="66e65-160">4 つのスライドを挿入する例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="66e65-161">配列内の各文字列は、プロパティに使用されるパターンの 1 つ以上に従う必要 `targetSlideId` があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="66e65-162">スライドは、配列に表示される順序に関係なく、元のプレゼンテーションに表示されるのと同じ相対順序で挿入されます。</span><span class="sxs-lookup"><span data-stu-id="66e65-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="66e65-163">ユーザーがソース プレゼンテーションのスライドの ID または作成 ID を検出できる実用的な方法はありません。</span><span class="sxs-lookup"><span data-stu-id="66e65-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="66e65-164">このため、コーディング時にソースの ID を知っている場合、またはアドインが実行時に一部のデータ ソースからそれらを取得できる場合にのみ、このプロパティを使用 `sourceSlideIds` できます。</span><span class="sxs-lookup"><span data-stu-id="66e65-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="66e65-165">ユーザーはスライド ID を記憶できないので、ユーザーがスライドを選択する方法 (タイトルやイメージなど) を有効にし、各タイトルまたはイメージをスライドの ID に関連付ける方法も必要です。</span><span class="sxs-lookup"><span data-stu-id="66e65-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="66e65-166">したがって、このプロパティは主にプレゼンテーション テンプレートのシナリオで使用されます。アドインは、挿入可能なスライドのプールとして機能する特定のプレゼンテーションのセットを操作するように `sourceSlideIds` 設計されています。</span><span class="sxs-lookup"><span data-stu-id="66e65-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="66e65-167">このようなシナリオでは、自分または顧客は、選択条件 (タイトルや画像など) と、考えられる一連のソース プレゼンテーションから構築されたスライドの ID またはスライド作成の ID を関連付けるデータ ソースを作成および維持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="66e65-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="66e65-168">スライドを削除する</span><span class="sxs-lookup"><span data-stu-id="66e65-168">Delete slides</span></span>

<span data-ttu-id="66e65-169">スライドを削除するには、スライドを表す [Slide](/javascript/api/powerpoint/powerpoint.slide) オブジェクトへの参照を取得し、メソッドを呼び出 `Slide.delete` します。</span><span class="sxs-lookup"><span data-stu-id="66e65-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="66e65-170">次に、4 番目のスライドを削除する例を示します。</span><span class="sxs-lookup"><span data-stu-id="66e65-170">The following is an example in which the 4th slide is deleted.</span></span>

```javascript
async function deleteSlide() {
  await PowerPoint.run(async function(context) {

    // The slide index is zero-based. 
    const slide = context.presentation.slides.getItemAt(3);
    slide.delete();
    await context.sync();
  });
}
```
