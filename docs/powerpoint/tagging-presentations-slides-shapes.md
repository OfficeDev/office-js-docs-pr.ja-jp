---
title: プレゼンテーション、スライド、図形にカスタム タグを使用PowerPoint
description: プレゼンテーション、スライド、図形に関するカスタム メタデータにタグを使用する方法について説明します。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 03f1656919ed16b801e97623f7f69c9f4adfaac8
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514210"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>プレゼンテーション、スライド、図形にカスタム タグを使用PowerPoint

アドインは、"tags" と呼ばれるキーと値のペアの形式で、スライド上のプレゼンテーション、特定のスライド、および特定の図形にカスタム メタデータを添付できます。

タグの使用には、主に次の 2 つのシナリオがあります。

- スライドまたは図形に適用すると、タグを使用すると、オブジェクトをバッチ処理用に分類できます。 たとえば、プレゼンテーションに、東地域のプレゼンテーションに含める必要があるスライドがいくつかあるとしますが、西側の領域には含めません。 同様に、西側にのみ表示する別のスライドがあります。 アドインは、キーと値を持つタグを作成し、東側でのみ使用する必要があるスライド `REGION` `East` に適用できます。 タグの値は、西地域にのみ表示する必要があるスライド `West` に対して設定されます。 東へのプレゼンテーションの直前に、アドインのボタンがコードを実行し、タグの値をチェックするスライドをループ `REGION` 処理します。 領域が削除された `West` スライド。 その後、ユーザーはアドインを閉じ、スライド ショーを開始します。
- プレゼンテーションに適用すると、実質的にタグはプレゼンテーション ドキュメント内のカスタム プロパティになります (Word の [CustomProperty](/javascript/api/word/word.customproperty) に似ています)。

## <a name="tag-slides-and-shapes"></a>スライドと図形にタグを付け

タグはキーと値のペアで、値は常に型であり `string` [、Tag](/javascript/api/powerpoint/powerpoint.tag) オブジェクトで表されます。 Presentation オブジェクト、Slide オブジェクト[](/javascript/api/powerpoint/powerpoint.slide)[、Shape](/javascript/api/powerpoint/powerpoint.presentation)オブジェクトなどの親オブジェクト[](/javascript/api/powerpoint/powerpoint.shape)の各種類には `tags` [、TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection)型のプロパティがあります。

### <a name="add-update-and-delete-tags"></a>タグの追加、更新、および削除

タグをオブジェクトに追加するには、親オブジェクトのプロパティ [の TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) メソッドを呼び出 `tags` します。 次のコードでは、プレゼンテーションの最初のスライドに 2 つのタグを追加します。 このコードについては、以下の点に注意してください。

- メソッドの最初のパラメーター `add` は、キーと値のペアのキーです。
- 2 番目のパラメーターは値です。
- キーは大文字です。 これはメソッドでは厳密に必須ではありませんが、キーは常に PowerPoint によって大文字として格納され、タグ関連のメソッドによってはキーを大文字で表す必要があります。そのため、タグ キーには常にコードで大文字を使用することをお勧めします。 `add` 

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

この `add` メソッドは、タグの更新にも使用されます。 次のコードは、タグの値を変更 `PLANET` します。

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

タグを削除するには、親オブジェクトのメソッドを呼び出し、タグのキーを `delete` `TagsCollection` パラメーターとして渡します。 例については、「プレゼンテーションでカスタム [メタデータを設定する」を参照してください](#set-custom-metadata-on-the-presentation)。

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>タグを使用してスライドと図形を選択的に処理する

次のシナリオを検討してください。 Contoso Consulting には、すべての新しい顧客に対して表示されるプレゼンテーションがあります。 ただし、一部のスライドは、"プレミアム" 状態の支払いを受け取ったユーザーにのみ表示する必要があります。 プレミアム以外のユーザーにプレゼンテーションを表示する前に、そのプレゼンテーションのコピーを作成し、プレミアムユーザーだけが表示するスライドを削除します。 アドインを使用すると、Contoso はプレミアムユーザー用のスライドにタグを付け、必要に応じてこれらのスライドを削除できます。 次の一覧では、この機能を作成するための主要なコーディング手順の概要を示します。

1. 現在選択されているスライドに顧客向けとしてタグ付けするメソッドを作成 `Premium` します。 このコードについては、以下の点に注意してください。

    - 関数 `getSelectedSlideIndex` は次の手順で定義されます。 現在選択されているスライドの 1 ベースのインデックスを返します。
    - `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)メソッドは 0 から始まないので、関数によって返される値をデクレメントする必要があります。

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

    - 共通 JavaScript API [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_)メソッドを使用します。
    - 呼び出 `getSelectedDataAsync` しは、promise-returning 関数に埋め込まれている。 これを行う理由と方法の詳細については [、「Promise-returning](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)関数で一般的な API をラップする」を参照してください。
    - `getSelectedDataAsync` 複数のスライドを選択できるので、配列を返します。 このシナリオでは、ユーザーが選択したスライドは 1 つだけなので、コードは最初の (0 番目) スライドを取得します。これが選択された唯一のスライドです。
    - スライドの値は、ユーザーが [UI サムネイル] ウィンドウのスライドの横に表示PowerPoint `index` 値です。

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

3. 次のコードは、プレミアムユーザーにタグ付けされたスライドを削除するメソッドを作成します。 このコードについては、以下の点に注意してください。

    - タグのプロパティとプロパティは、 の後に読み取りを行うので `key` `value` `context.sync` 、最初に読み込む必要があります。

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

## <a name="set-custom-metadata-on-the-presentation"></a>プレゼンテーションでカスタム メタデータを設定する

アドインは、プレゼンテーション全体にタグを適用することもできます。 これにより、Word での [CustomProperty](/javascript/api/word/word.customproperty)クラスの使用方法と同様に、ドキュメント レベルのメタデータにタグを使用できます。 ただし、Word クラス `CustomProperty` とは異なり、PowerPointの値は型のみです `string` 。

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

次のコードは、プレゼンテーションからタグを削除する例です。 タグのキーは親オブジェクトの `delete` メソッドに渡されます `TagsCollection` 。

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
