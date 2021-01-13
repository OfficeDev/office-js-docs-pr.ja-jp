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
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation"></a>PowerPoint プレゼンテーションのスライドの挿入と削除

PowerPoint アドインは、PowerPoint のアプリケーション固有の JavaScript ライブラリを使用して、1 つのプレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。 挿入したスライドで、元のプレゼンテーションの書式を維持するか、またはターゲット プレゼンテーションの書式設定を維持するかは制御できます。 プレゼンテーションからスライドを削除できます。

スライド挿入 API は、主にプレゼンテーション テンプレートのシナリオで使用されます。アドインによって挿入できるスライドのプールとして機能する、いくつかの既知のプレゼンテーションがあります。 このようなシナリオでは、自分または顧客のどちらかが、選択条件 (スライド タイトルや画像など) とスライドの ID を関連付けるデータ ソースを作成および維持する必要があります。 この API は、ユーザーが任意のプレゼンテーションからスライドを挿入できるシナリオでも使用できますが、そのシナリオでは、ユーザーは実質的にソース プレゼンテーションからすべてのスライドを挿入する方法に制限されます。 詳細 [については、「挿入するスライドの選択](#selecting-which-slides-to-insert) 」を参照してください。

1 つのプレゼンテーションから別のプレゼンテーションにスライドを挿入するには、2 つの手順があります。

1. ソース プレゼンテーション ファイル (.pptx) を base64 形式の文字列に変換します。
1. Base64 ファイルから現在のプレゼンテーションに 1 つ以上のスライドを挿入するには、この `insertSlidesFromBase64` メソッドを使用します。

## <a name="convert-the-source-presentation-to-base64"></a>ソース プレゼンテーションを base64 に変換する

ファイルを base64 に変換する方法は多数あります。 使用するプログラミング言語とライブラリ、およびアドインのサーバー側とクライアント側のどちらで変換するかは、シナリオによって決まります。 ほとんどの場合 [、FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側で JavaScript で変換を行います。 次の例は、この方法を示しています。

1. まず、ソース PowerPoint ファイルへの参照を取得します。 この例では、種類のコントロールを `<input>` 使用して、ユーザーにファイルの選択を求 `file` めるメッセージを表示します。 次のマークアップをアドイン ページに追加します。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    このマークアップは、次のスクリーンショットの UI をページに追加します。

    ![HTML ファイルの種類の入力コントロールの前に、「スライドを挿入する PowerPoint プレゼンテーションを選択する」という説明文が表示されているスクリーンショット コントロールは、"ファイルの選択" というラベルの付いたボタンと、その後に "No file chosen" という文で構成されます。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > PowerPoint ファイルを取得する方法は他に多数あります。 たとえば、ファイルが OneDrive または SharePoint に保存されている場合は、Microsoft Graph を使用してダウンロードできます。 詳細については [、「Microsoft Graph でのファイルの操作」および「Microsoft Graph](/graph/api/resources/onedrive) を使用した [ファイルへのアクセス」を参照してください](/learn/modules/msgraph-access-file-data/)。

2. 次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当 `change` てる。 (次の手順 `storeFileAsBase64` で関数を作成します。

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 次のコードを追加します。 このコードについては、次の点に注意してください。

    - この `reader.readAsDataURL` メソッドは、ファイルを base64 に変換し、プロパティに格納 `reader.result` します。 メソッドが完了すると、イベント ハンドラーが `onload` トリガーされます。
    - イベント ハンドラーは、エンコードされたファイルからメタデータをトリミングし、エンコードされた文字列をグローバル `onload` 変数に格納します。
    - base64 でエンコードされた文字列は、後の手順で作成する別の関数によって読み取りが行なうので、グローバルに格納されます。

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>insertSlidesFromBase64 を使用してスライドを挿入する

アドインは [、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) メソッドを使用して、別の PowerPoint プレゼンテーションのスライドを現在のプレゼンテーションに挿入します。 次に、ソース プレゼンテーションのすべてのスライドを現在のプレゼンテーションの先頭に挿入し、挿入したスライドにソース ファイルの書式を保持する簡単な例を示します。 これは、PowerPoint プレゼンテーション ファイルの base64 エンコード バージョンを保持するグローバル変数 `chosenFileBase64` です。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを 2 番目のパラメーターとして渡して、挿入結果の一部の側面 (スライドを挿入する場所、ソースまたはターゲットの書式設定を取得するかどうかを含む) を制御できます `insertSlidesFromBase64` 。 次に例を示します。 このコードについては、以下の点に注意してください。

- プロパティには `formatting` 、"UseDestinationTheme" と "KeepSourceFormatting" の 2 つの値を指定できます。 必要に応じて、列挙型 `InsertSlideFormatting` (例: ) を使用できます `PowerPoint.InsertSlideFormatting.useDestinationTheme` 。
- この関数は、プロパティで指定されたスライドの直後に、ソース プレゼンテーションのスライドを挿入 `targetSlideId` します。 このプロパティの値は **、*nnn*#**、* *#* mmm***、または **_nnn_ #* mmm***の 3 つの形式のいずれかの文字列です。ここで *、nnn* はスライドの ID (通常は 3 桁) で *、mmmmm は* スライドの作成 ID (通常は 9 桁) です。 たとえば、, `267#763315295` `267#` , and `#763315295` .

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

もちろん、通常、ターゲット スライドの ID または作成 ID はコーディング時にはわかりません。 多くの場合、アドインはユーザーにターゲット スライドの選択を求める場合があります。 次の手順は、現在選択されているスライドの ***nnn*#** ID を取得し、ターゲット スライドとして使用する方法を示しています。

1. 共通 JavaScript API の [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用して、現在選択されているスライドの ID を取得する関数を作成します。 次に例を示します。 呼び出しは `getSelectedDataAsync` Promise を返す関数に埋め込まれている点に注意してください。 これを行う理由と方法の詳細については、「Promise を返す関数で Common-APIs [をラップする」を参照してください](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。

 
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

1. メイン関数の[PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_)内で新しい関数を呼び出し、パラメーターのプロパティの値として返される ID ("#" 記号と連結) を渡します。 `targetSlideId` `InsertSlideOptions` 次に例を示します。

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

### <a name="selecting-which-slides-to-insert"></a>挿入するスライドの選択

[InsertSlideOptions パラメーター](/javascript/api/powerpoint/powerpoint.insertslideoptions)を使用して、ソース プレゼンテーションから挿入するスライドを制御することもできます。 これを行うには、元のプレゼンテーションのスライドのスライドの配列をプロパティに割り当 `sourceSlideIds` てる必要があります。 4 つのスライドを挿入する例を次に示します。 配列内の各文字列は、プロパティに使用されるパターンの 1 つ以上に従う必要 `targetSlideId` があります。

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
> スライドは、配列に表示される順序に関係なく、元のプレゼンテーションに表示されるのと同じ相対順序で挿入されます。

ユーザーがソース プレゼンテーションのスライドの ID または作成 ID を検出できる実用的な方法はありません。 このため、コーディング時にソースの ID を知っている場合、またはアドインが実行時に一部のデータ ソースからそれらを取得できる場合にのみ、このプロパティを使用 `sourceSlideIds` できます。 ユーザーはスライド ID を記憶できないので、ユーザーがスライドを選択する方法 (タイトルやイメージなど) を有効にし、各タイトルまたはイメージをスライドの ID に関連付ける方法も必要です。

したがって、このプロパティは主にプレゼンテーション テンプレートのシナリオで使用されます。アドインは、挿入可能なスライドのプールとして機能する特定のプレゼンテーションのセットを操作するように `sourceSlideIds` 設計されています。 このようなシナリオでは、自分または顧客は、選択条件 (タイトルや画像など) と、考えられる一連のソース プレゼンテーションから構築されたスライドの ID またはスライド作成の ID を関連付けるデータ ソースを作成および維持する必要があります。

## <a name="delete-slides"></a>スライドを削除する

スライドを削除するには、スライドを表す [Slide](/javascript/api/powerpoint/powerpoint.slide) オブジェクトへの参照を取得し、メソッドを呼び出 `Slide.delete` します。 次に、4 番目のスライドを削除する例を示します。

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
