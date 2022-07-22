---
title: PowerPoint のプレゼンテーション、スライド、図形にカスタム タグを使用する
description: プレゼンテーション、スライド、図形に関するカスタム メタデータにタグを使用する方法について説明します。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: a30beea56286437b1c69461534ca13912107cecf
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958903"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>PowerPoint でプレゼンテーション、スライド、図形にカスタム タグを使用する

アドインは、"タグ" と呼ばれるキーと値のペアの形式で、カスタム メタデータをプレゼンテーション、特定のスライド、およびスライド上の特定の図形にアタッチできます。

タグを使用するには、主に次の 2 つのシナリオがあります。

- スライドまたは図形に適用すると、タグを使用すると、オブジェクトをバッチ処理用に分類できます。 たとえば、プレゼンテーションに、西部リージョンではなく、東リージョンへのプレゼンテーションに含める必要があるスライドがあるとします。 同様に、西部にのみ表示する別のスライドがあります。 アドインは、キー `REGION` と値 `East` を含むタグを作成し、それを東部でのみ使用する必要があるスライドに適用できます。 タグの値は、西リージョンにのみ表示する必要があるスライドに対して設定 `West` されます。 東側へのプレゼンテーションの直前に、アドインのボタンによって、タグの値をチェックするすべてのスライドをループ処理するコードが `REGION` 実行されます。 リージョン `West` が削除されるスライド。 その後、ユーザーはアドインを閉じ、スライド ショーを開始します。
- プレゼンテーションに適用すると、タグは実質的にプレゼンテーション ドキュメント内のカスタム プロパティです (Word の [CustomProperty](/javascript/api/word/word.customproperty) に似ています)。

## <a name="tag-slides-and-shapes"></a>スライドと図形にタグを付け

タグはキーと値のペアで、値は常に型 `string` であり、 [Tag](/javascript/api/powerpoint/powerpoint.tag) オブジェクトによって表されます。 [Presentation](/javascript/api/powerpoint/powerpoint.presentation)、[Slide](/javascript/api/powerpoint/powerpoint.slide)、[Shape](/javascript/api/powerpoint/powerpoint.shape) オブジェクトなどの親オブジェクトの各型には`tags`、[TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection) 型のプロパティがあります。

### <a name="add-update-and-delete-tags"></a>タグの追加、更新、削除

オブジェクトにタグを追加するには、親オブジェクトのプロパティの [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) メソッドを呼び出します `tags` 。 次のコードは、プレゼンテーションの最初のスライドに 2 つのタグを追加します。 このコードについては、以下の点に注意してください。

- メソッドの最初の `add` パラメーターは、キーと値のペアのキーです。
- 2 番目のパラメーターは値です。
- キーは大文字です。 これはメソッドでは厳密に必須 `add` ではありませんが、キーは常に PowerPoint によって大文字として格納されます。 *タグ関連のメソッドによっては、キーを大文字で表す必要* があるため、タグ キーには常に大文字を使用することをお勧めします。

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

この `add` メソッドは、タグの更新にも使用されます。 次のコードは、タグの値を変更します `PLANET` 。

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

タグを削除するには、その親`TagsCollection`オブジェクトのメソッドを`delete`呼び出し、タグのキーをパラメーターとして渡します。 例については、「 [プレゼンテーションでカスタム メタデータを設定する」を](#set-custom-metadata-on-the-presentation)参照してください。

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>タグを使用してスライドと図形を選択的に処理する

次のシナリオを考慮してください。Contoso Consulting には、すべての新規顧客に表示されるプレゼンテーションがあります。 ただし、一部のスライドは、"Premium" 状態を支払った顧客にのみ表示する必要があります。 プレゼンテーションをプレミアム以外の顧客に表示する前に、そのプレゼンテーションのコピーを作成し、プレミアム顧客だけが表示するスライドを削除します。 アドインを使用すると、Contoso はプレミアムユーザー向けのスライドにタグを付け、必要に応じてこれらのスライドを削除できます。 次の一覧では、この機能を作成するための主要なコーディング手順の概要を示します。

1. 現在選択されているスライドを顧客向けにタグ付けする関数を `Premium` 作成します。 このコードについては、以下の点に注意してください。

    - 関数は `getSelectedSlideIndex` 、次の手順で定義します。 現在選択されているスライドの 1 から始まるインデックスを返します。
    - [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) メソッドは 0 から始まるため、関数によって`getSelectedSlideIndex`返される値を減らす必要があります。

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

2. 次のコードでは、選択したスライドのインデックスを取得するメソッドを作成します。 このコードについては、以下の点に注意してください。

    - Common JavaScript API の [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドを使用します。
    - 呼び出しは `getSelectedDataAsync` 、promise-returning 関数に埋め込まれます。 これを行う理由と方法の詳細については、「 [Promise を返す関数で共通 API をラップする」を](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)参照してください。
    - `getSelectedDataAsync` は、複数のスライドを選択できるため、配列を返します。 このシナリオでは、ユーザーが 1 つしか選択していないため、コードは最初の (0 番目の) スライドを取得します。これは選択された唯一のスライドです。
    - スライドの値は `index` 、PowerPoint UI サムネイル ウィンドウのスライドの横に表示される 1 から始まる値です。

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

3. 次のコードでは、Premium のお客様向けにタグ付けされたスライドを削除する関数を作成します。 このコードについては、以下の点に注意してください。

    - タグの `key` プロパティは `value` 後で読み取 `context.sync`られるので、最初に読み込む必要があります。

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

## <a name="set-custom-metadata-on-the-presentation"></a>プレゼンテーションにカスタム メタデータを設定する

アドインは、プレゼンテーション全体にタグを適用することもできます。 これにより、 [Word で CustomProperty](/javascript/api/word/word.customproperty)クラスを使用する場合と同様に、ドキュメント レベルのメタデータにタグを使用できます。 ただし、Word `CustomProperty` クラスとは異なり、PowerPoint タグの値は型 `string`のみです。

次のコードは、プレゼンテーションにタグを追加する例です。 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

次のコードは、プレゼンテーションからタグを削除する例です。 タグのキーは、親`TagsCollection`オブジェクトの`delete`メソッドに渡されることに注意してください。

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
