---
title: PowerPoint でスライドを追加および削除する
description: スライドを追加および削除し、新しいスライドのマスターとレイアウトを指定する方法について学習します。
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 5c1b9750acb905fd8e92484bb960c70ba39a7ca9
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613946"
---
# <a name="add-and-delete-slides-in-powerpoint-preview"></a><span data-ttu-id="2e56d-103">PowerPoint でスライドを追加および削除する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="2e56d-103">Add and delete slides in PowerPoint (preview)</span></span>

<span data-ttu-id="2e56d-104">PowerPoint アドインは、プレゼンテーションにスライドを追加し、必要に応じて、どのスライド マスター、およびマスター マスターのレイアウトを新しいスライドに使用して指定できます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="2e56d-105">アドインはスライドも削除できます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2e56d-106">スライドを追加する API はプレビュー中です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-106">The APIs for adding slides are in preview.</span></span> <span data-ttu-id="2e56d-107">開発環境またはテスト環境で実験してくださいが、実稼働アドインには追加しません。</span><span class="sxs-lookup"><span data-stu-id="2e56d-107">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span> <span data-ttu-id="2e56d-108">スライドを *削除するための* API がリリースされました。</span><span class="sxs-lookup"><span data-stu-id="2e56d-108">The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="2e56d-109">スライドを追加する API は、主に、プレゼンテーション内のスライド マスターとレイアウトの ID がコーディング時に知られているか、実行時にデータ ソースで見つかるシナリオで使用されます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-109">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="2e56d-110">このようなシナリオでは、選択基準 (スライド マスターやレイアウトの名前やイメージなど) とスライド マスターおよびレイアウトの ID を関連付けるデータ ソースを作成および管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e56d-110">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="2e56d-111">API は、ユーザーが既定のスライド マスターとマスター の既定のレイアウトを使用するスライドを挿入できるシナリオや、ユーザーが既存のスライドを選択して、同じスライド マスターとレイアウト (ただし、同じコンテンツではない) を持つ新しいスライドを作成できるシナリオでも使用できます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-111">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="2e56d-112">詳細 [については、「使用するスライド マスターとレイアウトの選択](#selecting-which-slide-master-and-layout-to-use) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e56d-112">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd"></a><span data-ttu-id="2e56d-113">SlideCollection.add を使用してスライドを追加する</span><span class="sxs-lookup"><span data-stu-id="2e56d-113">Add a slide with SlideCollection.add</span></span>

<span data-ttu-id="2e56d-114">[SlideCollection.add メソッドを使用してスライドを追加](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-114">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="2e56d-115">次に、プレゼンテーションの既定のスライド マスターを使用するスライドと、そのマスター の最初のレイアウトを追加する簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-115">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="2e56d-116">メソッドは、プレゼンテーションの最後に常に新しいスライドを追加します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-116">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="2e56d-117">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-117">The following is an example:</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="2e56d-118">使用するスライド マスターとレイアウトの選択</span><span class="sxs-lookup"><span data-stu-id="2e56d-118">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="2e56d-119">[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)パラメーターを使用して、新しいスライドに使用するスライド マスターと、マスター 内で使用するレイアウトを制御します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-119">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="2e56d-120">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-120">The following is an example.</span></span> <span data-ttu-id="2e56d-121">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2e56d-121">Note the following about this code:</span></span>

- <span data-ttu-id="2e56d-122">オブジェクトのプロパティのどちらかまたは両方を含 `AddSlideOptions` めることができます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-122">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="2e56d-123">両方のプロパティを使用する場合は、指定したレイアウトが指定したマスターに属している必要があります。またはエラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-123">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="2e56d-124">プロパティが存在しない場合 (または値が空の文字列である場合)、既定のスライド マスターが使用され、そのスライド マスターのレイアウト `masterId` `layoutId` である必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e56d-124">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="2e56d-125">既定のスライド マスターは、プレゼンテーションの最後のスライドで使用されるスライド マスターです。</span><span class="sxs-lookup"><span data-stu-id="2e56d-125">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="2e56d-126">(プレゼンテーションに現在スライドがない場合、既定のスライド マスターはプレゼンテーションの最初のスライド マスターです。</span><span class="sxs-lookup"><span data-stu-id="2e56d-126">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="2e56d-127">プロパティが存在しない場合 (または値が空の文字列である) 場合は、the で指定されたマスター の最初の `layoutId` レイアウトが `masterId` 使用されます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-127">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="2e56d-128">どちらのプロパティも **、*nnnnnn*#** *#* 、\* mmmmmmm***、または \* nn mmmmm***のいずれかの文字列で *_、nnnn_ #\*\*は* マスターまたはレイアウトの ID (通常は 10 桁) で *、mmmmmmmmm は* マスターまたはレイアウトの作成 ID (通常は 6 ~ 10 桁) です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-128">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="2e56d-129">いくつかの例は `2147483690#2908289500` 、、 `2147483690#` 、および `#2908289500` です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-129">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

<span data-ttu-id="2e56d-130">ユーザーがスライド マスターまたはレイアウトの ID または作成 ID を検出できる実用的な方法はありません。</span><span class="sxs-lookup"><span data-stu-id="2e56d-130">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="2e56d-131">このため、コーディング時に ID を知っている場合、またはアドインが実行時に検出できる場合にのみ、このパラメーターを `AddSlideOptions` 使用できます。</span><span class="sxs-lookup"><span data-stu-id="2e56d-131">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="2e56d-132">ユーザーが ID を記憶する必要が生じないので、ユーザーがスライド (名前または画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-132">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="2e56d-133">したがって、このパラメーターは主に、アドインが特定のスライド マスターとレイアウトのセットで動作するように設計されたシナリオで使用 `AddSlideOptions` されます。その ID は既知です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-133">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="2e56d-134">このようなシナリオでは、ユーザーまたは顧客のどちらかが、選択基準 (スライド マスター名やレイアウト名やイメージなど) と対応する ID または作成用の ID を関連付けるデータ ソースを作成および管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e56d-134">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="2e56d-135">ユーザーに一致するスライドを選択する</span><span class="sxs-lookup"><span data-stu-id="2e56d-135">Have the user choose a matching slide</span></span>

<span data-ttu-id="2e56d-136">新しいスライドで既存のスライドで使用されるスライド マスターとレイアウトの同じ組み合わせを使用するシナリオでアドインを使用できる場合は、(1) ユーザーにスライドの選択を求めるプロンプトを表示し、(2) スライド マスターとレイアウトの ID を読み取る必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e56d-136">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="2e56d-137">次の手順では、一致するマスターとレイアウトを持つスライドを読み取り、スライドを追加する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-137">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="2e56d-138">選択したスライドのインデックスを取得するメソッドを作成します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-138">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="2e56d-139">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-139">The following is an example.</span></span> <span data-ttu-id="2e56d-140">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="2e56d-140">Note about this code:</span></span>

    - <span data-ttu-id="2e56d-141">共通 JavaScript API [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-141">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="2e56d-142">呼び出 `getSelectedDataAsync` しは Promise 戻り関数に埋め込まれている。</span><span class="sxs-lookup"><span data-stu-id="2e56d-142">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="2e56d-143">これを行う理由と方法の詳細については [、「Promise-returning](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)関数で一般的な API をラップする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e56d-143">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="2e56d-144">`getSelectedDataAsync` 複数のスライドを選択できるので、配列を返します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-144">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="2e56d-145">このシナリオでは、ユーザーが選択したスライドは 1 つだけなので、コードは最初の (0 番目) スライドを取得します。これが選択された唯一のスライドです。</span><span class="sxs-lookup"><span data-stu-id="2e56d-145">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="2e56d-146">スライドの値は、ユーザーがサムネイル ウィンドウのスライドの横に表示する `index` 1 ベースの値です。</span><span class="sxs-lookup"><span data-stu-id="2e56d-146">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. <span data-ttu-id="2e56d-147">スライドを追加するメイン関数 [の PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-147">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="2e56d-148">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-148">The following is an example:</span></span>

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a><span data-ttu-id="2e56d-149">スライドを削除する</span><span class="sxs-lookup"><span data-stu-id="2e56d-149">Delete slides</span></span>

<span data-ttu-id="2e56d-150">スライドを表す [Slide](/javascript/api/powerpoint/powerpoint.slide) オブジェクトへの参照を取得してスライドを削除し、メソッドを呼び出 `Slide.delete` します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-150">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="2e56d-151">4 番目のスライドを削除する例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2e56d-151">The following is an example in which the 4th slide is deleted:</span></span>

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
