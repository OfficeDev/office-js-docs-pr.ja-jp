---
title: PowerPoint プレゼンテーションでスライドを挿入および削除する
description: プレゼンテーションのスライドを別のプレゼンテーションに挿入する方法と、スライドを削除する方法について説明します。
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613705"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a><span data-ttu-id="e7535-103">PowerPoint プレゼンテーションでスライドを挿入および削除する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e7535-103">Insert and delete slides in a PowerPoint presentation (preview)</span></span>

<span data-ttu-id="e7535-104">PowerPoint アドインでは、PowerPoint のアプリケーション固有の JavaScript ライブラリを使用して、プレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="e7535-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="e7535-105">挿入したスライドに、元のプレゼンテーションの書式設定を保持するか、または対象となるプレゼンテーションの書式設定を保持するかを制御できます。</span><span class="sxs-lookup"><span data-stu-id="e7535-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="e7535-106">プレゼンテーションからスライドを削除することもできます。</span><span class="sxs-lookup"><span data-stu-id="e7535-106">You can also delete slides from the presentation.</span></span>

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

<span data-ttu-id="e7535-107">スライド挿入 Api は、主にプレゼンテーションテンプレートのシナリオで使用されます。これは、アドインによって挿入できるスライドのプールとして機能する既知のプレゼンテーションの数が少ないことです。</span><span class="sxs-lookup"><span data-stu-id="e7535-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="e7535-108">このようなシナリオでは、ユーザーまたは顧客は、選択基準 (スライドタイトル、画像など) とスライド Id を関連付けたデータソースを作成して維持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7535-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="e7535-109">この Api は、ユーザーが任意のプレゼンテーションからスライドを挿入できる場合にも使用できますが、このシナリオでは、ユーザーは、元のプレゼンテーションの *すべて* のスライドを挿入することに効果的に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e7535-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="e7535-110">詳細については、「 [挿入するスライドを選択](#selecting-which-slides-to-insert) する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="e7535-111">プレゼンテーションのスライドを別のプレゼンテーションに挿入するには、2つの手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="e7535-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="e7535-112">ソースプレゼンテーションファイル (.pptx) を base64 形式の文字列に変換します。</span><span class="sxs-lookup"><span data-stu-id="e7535-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="e7535-113">メソッドを使用して、 `insertSlidesFromBase64` base64 ファイルから現在のプレゼンテーションに1つまたは複数のスライドを挿入します。</span><span class="sxs-lookup"><span data-stu-id="e7535-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="e7535-114">ソースプレゼンテーションを base64 に変換する</span><span class="sxs-lookup"><span data-stu-id="e7535-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="e7535-115">ファイルを base64 に変換するには、さまざまな方法があります。</span><span class="sxs-lookup"><span data-stu-id="e7535-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="e7535-116">使用するプログラミング言語とライブラリ、およびアドインまたはクライアント側のサーバー側で変換するかどうかは、シナリオによって決まります。</span><span class="sxs-lookup"><span data-stu-id="e7535-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="e7535-117">通常、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側の JavaScript で変換を行います。</span><span class="sxs-lookup"><span data-stu-id="e7535-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="e7535-118">次の例は、この方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e7535-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="e7535-119">最初に、ソースの PowerPoint ファイルへの参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7535-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="e7535-120">この例では、 `<input>` 種類のコントロールを使用し `file` て、ユーザーにファイルの選択を求めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="e7535-121">次のマークアップをアドインページに追加します。</span><span class="sxs-lookup"><span data-stu-id="e7535-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="e7535-122">このマークアップは、次のスクリーンショットの UI をページに追加します。</span><span class="sxs-lookup"><span data-stu-id="e7535-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![HTML ファイルの種類の入力コントロールの前に説明文を表示するスクリーンショット。「スライドを挿入する PowerPoint プレゼンテーションを選択する」を参照してください。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="e7535-125">PowerPoint ファイルを取得する方法は他にもたくさんあります。</span><span class="sxs-lookup"><span data-stu-id="e7535-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="e7535-126">たとえば、ファイルが OneDrive または SharePoint に保存されている場合は、Microsoft Graph を使用してダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="e7535-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="e7535-127">詳細については、「 [Microsoft graph でファイルを処理](/graph/api/resources/onedrive) する」および「 [Microsoft graph でファイルにアクセス](/learn/modules/msgraph-access-file-data/)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="e7535-128">次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当て `change` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="e7535-129">(この関数は、 `storeFileAsBase64` 次の手順で作成します)。</span><span class="sxs-lookup"><span data-stu-id="e7535-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="e7535-130">次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e7535-130">Add the following code.</span></span> <span data-ttu-id="e7535-131">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="e7535-132">`reader.readAsDataURL`このメソッドは、ファイルを base64 に変換して、プロパティに格納し `reader.result` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="e7535-133">メソッドが完了すると、イベントハンドラーがトリガーされ `onload` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="e7535-134">イベントハンドラーは、エンコードされた `onload` ファイルからメタデータをトリミングし、エンコードされた文字列をグローバル変数に格納します。</span><span class="sxs-lookup"><span data-stu-id="e7535-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="e7535-135">Base64 でエンコードされた文字列は、後の手順で作成した別の関数によって読み取られるため、グローバルに格納されます。</span><span class="sxs-lookup"><span data-stu-id="e7535-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="e7535-136">InsertSlidesFromBase64 を使用してスライドを挿入する</span><span class="sxs-lookup"><span data-stu-id="e7535-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="e7535-137">アドインは、 [insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) メソッドを使用して、別の PowerPoint プレゼンテーションのスライドを現在のプレゼンテーションに挿入します。</span><span class="sxs-lookup"><span data-stu-id="e7535-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="e7535-138">次の簡単な例は、元のプレゼンテーションのすべてのスライドを現在のプレゼンテーションの先頭に挿入し、挿入したスライドにソースファイルの書式を保持します。</span><span class="sxs-lookup"><span data-stu-id="e7535-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="e7535-139">これ `chosenFileBase64` は、PowerPoint プレゼンテーションファイルの base64 でエンコードされたバージョンを保持するグローバル変数であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="e7535-140">[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを2番目のパラメーターとして渡すことによって、挿入結果のいくつかの側面を制御できます。これには、スライドが挿入される場所や、ソースまたはターゲットの書式設定を取得するかどうかなどがあり `insertSlidesFromBase64` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="e7535-141">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-141">The following is an example.</span></span> <span data-ttu-id="e7535-142">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-142">About this code, note:</span></span>

- <span data-ttu-id="e7535-143">プロパティには、 `formatting` "UseDestinationTheme" と "KeepSourceFormatting" という2つの値があります。</span><span class="sxs-lookup"><span data-stu-id="e7535-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="e7535-144">必要に応じて、 `InsertSlideFormatting` enum (など) を使用でき `PowerPoint.InsertSlideFormatting.useDestinationTheme` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="e7535-145">関数は、プロパティで指定されたスライドの直後に、元のプレゼンテーションのスライドを挿入し `targetSlideId` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="e7535-146">このプロパティの値は、次の3つの形式のうちの1つです。 \***nnn \* #**、\* *#* mmmmmmmmm \* \* \*、または \**_nnn_ #* mmmmmmmmm \* \* \*、 *nnn* はスライドの id (通常は3桁)、 *mmmmmmmmm* はスライドの作成 ID (通常は9桁) です。</span><span class="sxs-lookup"><span data-stu-id="e7535-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="e7535-147">例として、、、などがあり `267#763315295` `267#` `#763315295` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="e7535-148">当然のことですが、通常は、ターゲットスライドの ID または作成 ID のコーディング時にはわかりません。</span><span class="sxs-lookup"><span data-stu-id="e7535-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="e7535-149">一般的に、アドインは、ターゲットスライドを選択するようにユーザーに要求します。</span><span class="sxs-lookup"><span data-stu-id="e7535-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="e7535-150">次の手順は、現在選択されているスライドの \***nnn \* #** ID を取得し、それをターゲットスライドとして使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e7535-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="e7535-151">共通 JavaScript Api の [Office.context.document](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) を使用して、現在選択されているスライドの ID を取得する関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e7535-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="e7535-152">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-152">The following is an example.</span></span> <span data-ttu-id="e7535-153">の呼び出し `getSelectedDataAsync` は、Promise を返す関数に埋め込まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="e7535-154">その理由と方法の詳細については、「 [関数を返す関数のラップ Common-APIs](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7535-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="e7535-155">Main 関数の [PowerPoint. run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出し、返される ID ("#" 記号で連結される) を `targetSlideId` パラメーターのプロパティの値として渡します `InsertSlideOptions` 。</span><span class="sxs-lookup"><span data-stu-id="e7535-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="e7535-156">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="e7535-157">挿入するスライドの選択</span><span class="sxs-lookup"><span data-stu-id="e7535-157">Selecting which slides to insert</span></span>

<span data-ttu-id="e7535-158">[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)パラメーターを使用して、ソースプレゼンテーションから挿入するスライドを制御することもできます。</span><span class="sxs-lookup"><span data-stu-id="e7535-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="e7535-159">これを行うには、移動元のプレゼンテーションのスライド Id の配列を `sourceSlideIds` プロパティに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="e7535-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="e7535-160">次に、4つのスライドを挿入する例を示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="e7535-161">配列内の各文字列は、プロパティに使用されている1つまたは複数のパターンに従う必要があることに注意してください `targetSlideId` 。</span><span class="sxs-lookup"><span data-stu-id="e7535-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="e7535-162">スライドは、配列に表示される順序に関係なく、ソースプレゼンテーションに表示される相対的な順序で挿入されます。</span><span class="sxs-lookup"><span data-stu-id="e7535-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="e7535-163">ユーザーがソースプレゼンテーションのスライドの ID または作成 ID を検出するための実用的な方法はありません。</span><span class="sxs-lookup"><span data-stu-id="e7535-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="e7535-164">このため、このプロパティは、 `sourceSlideIds` コーディング時にソース id がわかっている場合、またはデータソースから実行時にアドインで取得できる場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="e7535-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="e7535-165">ユーザーがスライド Id を記憶することは予想できないため、ユーザーがスライドを選択できるようにするには、タイトルや画像によるスライドの選択が可能になり、各タイトルまたは画像をスライドの ID と関連付けることが必要になります。</span><span class="sxs-lookup"><span data-stu-id="e7535-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="e7535-166">そのため、この `sourceSlideIds` プロパティは、主にプレゼンテーションテンプレートのシナリオで使用されます。アドインは、挿入可能なスライドのプールとして機能する特定のプレゼンテーションセットで機能するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="e7535-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="e7535-167">このようなシナリオでは、お客様またはお客様は、選択基準 (タイトル、画像など) を含むデータソースを作成して、使用可能なソースプレゼンテーションセットから構成されているスライド作成 Id を関連付けて保持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7535-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="e7535-168">スライドの削除</span><span class="sxs-lookup"><span data-stu-id="e7535-168">Delete slides</span></span>

<span data-ttu-id="e7535-169">スライドを表す [slide オブジェクトへの参照](/javascript/api/powerpoint/powerpoint.slide) を取得して、メソッドを呼び出すことで、スライドを削除でき `Slide.delete` ます。</span><span class="sxs-lookup"><span data-stu-id="e7535-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="e7535-170">次に、4番目のスライドを削除する例を示します。</span><span class="sxs-lookup"><span data-stu-id="e7535-170">The following is an example in which the 4th slide is deleted.</span></span>

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
