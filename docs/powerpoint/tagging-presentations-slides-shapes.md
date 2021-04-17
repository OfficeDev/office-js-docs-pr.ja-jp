---
title: PowerPoint でプレゼンテーション、スライド、図形にカスタム タグを使用する
description: プレゼンテーション、スライド、図形に関するカスタム メタデータにタグを使用する方法について説明します。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fbb13e67da1f7962fc2c0b8d45689f259b015014
ms.sourcegitcommit: 58d394fa49308ecf93cd53f7d3fb6e316ff56209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/16/2021
ms.locfileid: "51876861"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a><span data-ttu-id="25098-103">PowerPoint でプレゼンテーション、スライド、図形にカスタム タグを使用する</span><span class="sxs-lookup"><span data-stu-id="25098-103">Use custom tags for presentations, slides, and shapes in PowerPoint</span></span>

<span data-ttu-id="25098-104">アドインは、"tags" と呼ばれるキーと値のペアの形式で、スライド上のプレゼンテーション、特定のスライド、および特定の図形にカスタム メタデータを添付できます。</span><span class="sxs-lookup"><span data-stu-id="25098-104">An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="25098-105">タグの API はプレビュー中です。</span><span class="sxs-lookup"><span data-stu-id="25098-105">The APIs for tags are in preview.</span></span> <span data-ttu-id="25098-106">開発環境またはテスト環境で実験してくださいが、実稼働アドインには追加しません。</span><span class="sxs-lookup"><span data-stu-id="25098-106">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>

<span data-ttu-id="25098-107">タグの使用には、主に次の 2 つのシナリオがあります。</span><span class="sxs-lookup"><span data-stu-id="25098-107">There are two main scenarios for using tags:</span></span>

- <span data-ttu-id="25098-108">スライドまたは図形に適用すると、タグを使用すると、オブジェクトをバッチ処理用に分類できます。</span><span class="sxs-lookup"><span data-stu-id="25098-108">When applied to a slide or a shape, a tag enables the object to be categorized for batch processing.</span></span> <span data-ttu-id="25098-109">たとえば、プレゼンテーションに、東地域のプレゼンテーションに含める必要があるスライドがいくつかあるとしますが、西側の領域には含めません。</span><span class="sxs-lookup"><span data-stu-id="25098-109">For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region.</span></span> <span data-ttu-id="25098-110">同様に、西側にのみ表示する別のスライドがあります。</span><span class="sxs-lookup"><span data-stu-id="25098-110">Similarly, there are alternative slides that should be shown only to the West.</span></span> <span data-ttu-id="25098-111">アドインは、キーと値を持つタグを作成し、東側でのみ使用する必要があるスライド `REGION` `East` に適用できます。</span><span class="sxs-lookup"><span data-stu-id="25098-111">Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East.</span></span> <span data-ttu-id="25098-112">タグの値は、西地域にのみ表示する必要があるスライド `West` に対して設定されます。</span><span class="sxs-lookup"><span data-stu-id="25098-112">The tag's value is set to `West` for the slides that should only be shown to the West region.</span></span> <span data-ttu-id="25098-113">東へのプレゼンテーションの直前に、アドインのボタンがコードを実行し、タグの値をチェックするスライドをループ `REGION` 処理します。</span><span class="sxs-lookup"><span data-stu-id="25098-113">Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag.</span></span> <span data-ttu-id="25098-114">領域が削除された `West` スライド。</span><span class="sxs-lookup"><span data-stu-id="25098-114">Slides where the region is `West` are deleted.</span></span> <span data-ttu-id="25098-115">その後、ユーザーはアドインを閉じ、スライド ショーを開始します。</span><span class="sxs-lookup"><span data-stu-id="25098-115">The user then closes the add-in and starts the slide show.</span></span>
- <span data-ttu-id="25098-116">プレゼンテーションに適用すると、実質的にタグはプレゼンテーション ドキュメント内のカスタム プロパティになります (Word の [CustomProperty](/javascript/api/word/word.customproperty) に似ています)。</span><span class="sxs-lookup"><span data-stu-id="25098-116">When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).</span></span>

## <a name="tag-slides-and-shapes"></a><span data-ttu-id="25098-117">スライドと図形にタグを付け</span><span class="sxs-lookup"><span data-stu-id="25098-117">Tag slides and shapes</span></span>

<span data-ttu-id="25098-118">タグはキーと値のペアで、値は常に型であり `string` [、Tag](/javascript/api/powerpoint/powerpoint.tag) オブジェクトで表されます。</span><span class="sxs-lookup"><span data-stu-id="25098-118">A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object.</span></span> <span data-ttu-id="25098-119">Presentation オブジェクト、Slide オブジェクト[](/javascript/api/powerpoint/powerpoint.slide)[、Shape](/javascript/api/powerpoint/powerpoint.presentation)オブジェクトなどの親オブジェクト[](/javascript/api/powerpoint/powerpoint.shape)の各種類には `tags` [、TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection)型のプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="25098-119">Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span></span>

### <a name="add-update-and-delete-tags"></a><span data-ttu-id="25098-120">タグの追加、更新、および削除</span><span class="sxs-lookup"><span data-stu-id="25098-120">Add, update, and delete tags</span></span>

<span data-ttu-id="25098-121">タグをオブジェクトに追加するには、親オブジェクトのプロパティ [の TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) メソッドを呼び出 `tags` します。</span><span class="sxs-lookup"><span data-stu-id="25098-121">To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) method of the parent object's `tags` property.</span></span> <span data-ttu-id="25098-122">次のコードでは、プレゼンテーションの最初のスライドに 2 つのタグを追加します。</span><span class="sxs-lookup"><span data-stu-id="25098-122">The following code adds two tags to the first slide of a presentation.</span></span> <span data-ttu-id="25098-123">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25098-123">About this code, note:</span></span>

- <span data-ttu-id="25098-124">メソッドの最初のパラメーター `add` は、キーと値のペアのキーです。</span><span class="sxs-lookup"><span data-stu-id="25098-124">The first parameter of the `add` method is the key in the key-value pair.</span></span> 
- <span data-ttu-id="25098-125">2 番目のパラメーターは値です。</span><span class="sxs-lookup"><span data-stu-id="25098-125">The second parameter is the value.</span></span>
- <span data-ttu-id="25098-126">キーは大文字です。</span><span class="sxs-lookup"><span data-stu-id="25098-126">The key is in uppercase letters.</span></span> <span data-ttu-id="25098-127">これはメソッドでは厳密には必須ではありませんが、キーは常に PowerPoint によって大文字として格納され、タグ関連のメソッドによっては、キーを大文字で表す必要があります。そのため、タグ キーのコードでは常に大文字を使用することをお勧めします。 `add` </span><span class="sxs-lookup"><span data-stu-id="25098-127">This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.</span></span>

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

<span data-ttu-id="25098-128">この `add` メソッドは、タグの更新にも使用されます。</span><span class="sxs-lookup"><span data-stu-id="25098-128">The `add` method is also used to update a tag.</span></span> <span data-ttu-id="25098-129">次のコードは、タグの値を変更 `PLANET` します。</span><span class="sxs-lookup"><span data-stu-id="25098-129">The following code changes the value of the `PLANET` tag.</span></span>

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

<span data-ttu-id="25098-130">タグを削除するには、親オブジェクトのメソッドを呼び出し、タグのキーを `delete` `TagsCollection` パラメーターとして渡します。</span><span class="sxs-lookup"><span data-stu-id="25098-130">To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter.</span></span> <span data-ttu-id="25098-131">例については、「プレゼンテーションでカスタム [メタデータを設定する」を参照してください](#set-custom-metadata-on-the-presentation)。</span><span class="sxs-lookup"><span data-stu-id="25098-131">For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).</span></span>

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a><span data-ttu-id="25098-132">タグを使用してスライドと図形を選択的に処理する</span><span class="sxs-lookup"><span data-stu-id="25098-132">Use tags to selectively process slides and shapes</span></span>

<span data-ttu-id="25098-133">次のシナリオを検討してください。 Contoso Consulting には、すべての新しい顧客に対して表示されるプレゼンテーションがあります。</span><span class="sxs-lookup"><span data-stu-id="25098-133">Consider the following scenario: Contoso Consulting has a presentation they show to all new customers.</span></span> <span data-ttu-id="25098-134">ただし、一部のスライドは、"プレミアム" 状態の支払いを受け取ったユーザーにのみ表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="25098-134">But some slides should only be shown to customers that have paid for "premium" status.</span></span> <span data-ttu-id="25098-135">プレミアム以外のユーザーにプレゼンテーションを表示する前に、そのプレゼンテーションのコピーを作成し、プレミアムユーザーだけが表示するスライドを削除します。</span><span class="sxs-lookup"><span data-stu-id="25098-135">Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see.</span></span> <span data-ttu-id="25098-136">アドインを使用すると、Contoso はプレミアムユーザー用のスライドにタグを付け、必要に応じてこれらのスライドを削除できます。</span><span class="sxs-lookup"><span data-stu-id="25098-136">An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed.</span></span> <span data-ttu-id="25098-137">次の一覧では、この機能を作成するための主要なコーディング手順の概要を示します。</span><span class="sxs-lookup"><span data-stu-id="25098-137">The following list outlines the major coding steps to create this functionality.</span></span>

1. <span data-ttu-id="25098-138">現在選択されているスライドに顧客向けとしてタグ付けするメソッドを作成 `Premium` します。</span><span class="sxs-lookup"><span data-stu-id="25098-138">Create a method that tags the currently selected slide as intended for `Premium` customers.</span></span> <span data-ttu-id="25098-139">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25098-139">About this code, note:</span></span>

    - <span data-ttu-id="25098-140">関数 `getSelectedSlideIndex` は次の手順で定義されます。</span><span class="sxs-lookup"><span data-stu-id="25098-140">The `getSelectedSlideIndex` function is defined in the next step.</span></span> <span data-ttu-id="25098-141">現在選択されているスライドの 1 ベースのインデックスを返します。</span><span class="sxs-lookup"><span data-stu-id="25098-141">It returns the 1-based index of the currently selected slide.</span></span>
    - <span data-ttu-id="25098-142">`getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)メソッドは 0 から始まないので、関数によって返される値をデクレメントする必要があります。</span><span class="sxs-lookup"><span data-stu-id="25098-142">The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) method is 0-based.</span></span>

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. <span data-ttu-id="25098-143">次のコードでは、選択したスライドのインデックスを取得するメソッドを作成します。</span><span class="sxs-lookup"><span data-stu-id="25098-143">The following code creates a method to get the index of the selected slide.</span></span> <span data-ttu-id="25098-144">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25098-144">About this code, note:</span></span>

    - <span data-ttu-id="25098-145">共通 JavaScript API [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="25098-145">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="25098-146">呼び出 `getSelectedDataAsync` しは、promise-returning 関数に埋め込まれている。</span><span class="sxs-lookup"><span data-stu-id="25098-146">The call to `getSelectedDataAsync` is embedded in a promise-returning function.</span></span> <span data-ttu-id="25098-147">これを行う理由と方法の詳細については [、「Promise-returning](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)関数で一般的な API をラップする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25098-147">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="25098-148">`getSelectedDataAsync` 複数のスライドを選択できるので、配列を返します。</span><span class="sxs-lookup"><span data-stu-id="25098-148">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="25098-149">このシナリオでは、ユーザーが選択したスライドは 1 つだけなので、コードは最初の (0 番目) スライドを取得します。これが選択された唯一のスライドです。</span><span class="sxs-lookup"><span data-stu-id="25098-149">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="25098-150">スライドの値は、ユーザーが PowerPoint UI サムネイル ウィンドウのスライドの横に表示する `index` 1 ベースの値です。</span><span class="sxs-lookup"><span data-stu-id="25098-150">The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.</span></span>

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

3. <span data-ttu-id="25098-151">次のコードは、プレミアムユーザーにタグ付けされたスライドを削除するメソッドを作成します。</span><span class="sxs-lookup"><span data-stu-id="25098-151">The following code creates a method to delete slides that are tagged for premium customers.</span></span> <span data-ttu-id="25098-152">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25098-152">About this code, note:</span></span>

    - <span data-ttu-id="25098-153">タグのプロパティとプロパティは、 の後に読み取りを行うので `key` `value` `context.sync` 、最初に読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="25098-153">Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.</span></span>

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a><span data-ttu-id="25098-154">プレゼンテーションでカスタム メタデータを設定する</span><span class="sxs-lookup"><span data-stu-id="25098-154">Set custom metadata on the presentation</span></span>

<span data-ttu-id="25098-155">アドインは、プレゼンテーション全体にタグを適用することもできます。</span><span class="sxs-lookup"><span data-stu-id="25098-155">Add-ins can also apply tags to the presentation as a whole.</span></span> <span data-ttu-id="25098-156">これにより、Word での [CustomProperty](/javascript/api/word/word.customproperty)クラスの使用方法と同様に、ドキュメント レベルのメタデータにタグを使用できます。</span><span class="sxs-lookup"><span data-stu-id="25098-156">This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word.</span></span> <span data-ttu-id="25098-157">ただし、Word クラス `CustomProperty` とは異なり、PowerPoint タグの値は型のみです `string` 。</span><span class="sxs-lookup"><span data-stu-id="25098-157">But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.</span></span>

<span data-ttu-id="25098-158">次のコードは、プレゼンテーションにタグを追加する例です。</span><span class="sxs-lookup"><span data-stu-id="25098-158">The following code is an example of adding a tag to a presentation.</span></span> 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

<span data-ttu-id="25098-159">次のコードは、プレゼンテーションからタグを削除する例です。</span><span class="sxs-lookup"><span data-stu-id="25098-159">The following code is an example of deleting a tag from a presentation.</span></span> <span data-ttu-id="25098-160">タグのキーは親オブジェクトの `delete` メソッドに渡されます `TagsCollection` 。</span><span class="sxs-lookup"><span data-stu-id="25098-160">Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.</span></span>

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
