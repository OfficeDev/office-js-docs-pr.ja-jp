---
title: プレゼンテーションにスライドをPowerPointする
description: プレゼンテーションから別のプレゼンテーションにスライドを挿入する方法について説明します。
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: d9c50b87e7ba702a2cffcef5ca94dfb0d39b1af0
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671766"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>プレゼンテーションにスライドをPowerPointする

1 PowerPoint PowerPointアドインは、アプリケーション固有の JavaScript ライブラリを使用して、1 つのプレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。 挿入されたスライドがソース プレゼンテーションの書式設定を保持するか、ターゲット プレゼンテーションの書式設定を保持するかどうかを制御できます。

スライド挿入 API は、主にプレゼンテーション テンプレートのシナリオで使用されます。既知のプレゼンテーションは、アドインによって挿入できるスライドのプールとして機能します。 このようなシナリオでは、ユーザーまたは顧客のどちらかが、スライドのタイトルや画像などの選択基準とスライドの ID を関連付けるデータ ソースを作成および管理する必要があります。 API は、ユーザーが任意のプレゼンテーションからスライドを挿入できるシナリオでも使用できますが、そのシナリオでは、ユーザーは実質的にソース プレゼンテーションからすべてのスライドを挿入する制限があります。 詳細 [については、「挿入するスライドの選択](#selecting-which-slides-to-insert) 」を参照してください。

プレゼンテーションから別のプレゼンテーションにスライドを挿入するには、2 つの手順があります。

1. ソース プレゼンテーション ファイル (.pptx) を base64 形式の文字列に変換します。
1. base64 ファイルから現在のプレゼンテーションに 1 つ以上のスライドを挿入するには、この `insertSlidesFromBase64` メソッドを使用します。

## <a name="convert-the-source-presentation-to-base64"></a>ソース プレゼンテーションを base64 に変換する

ファイルを base64 に変換する方法は多数あります。 使用するプログラミング言語とライブラリ、およびアドインのサーバー側またはクライアント側で変換するかどうかは、シナリオによって決まります。 最も一般的には [、FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側の JavaScript で変換を行います。 次の例は、このプラクティスを示しています。

1. まず、ソース ファイルへの参照を取得PowerPointします。 この例では、種類のコントロールを `<input>` 使用 `file` して、ユーザーにファイルの選択を求めるメッセージを表示します。 アドイン ページに次のマークアップを追加します。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    このマークアップは、次のスクリーンショットの UI をページに追加します。

    ![HTML ファイルの種類の入力コントロールの前に「スライドを挿入するプレゼンテーションを選択する」というPowerPointを示すスクリーンショット。 コントロールは、"ファイルの選択" というラベルの付いたボタンの後に"ファイルが選択されません" という文で構成されます。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > ファイルを取得する方法は他にPowerPointがあります。 たとえば、ファイルがサーバーまたはサーバーに保存されているOneDrive SharePoint、Microsoft Graphを使用してダウンロードできます。 詳細については[、「Microsoft Graph](/graph/api/resources/onedrive)ファイルの操作」および「Access Files with [Microsoft Graph」を参照してください](/learn/modules/msgraph-access-file-data/)。

2. 次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当 `change` てる。 (次の手順 `storeFileAsBase64` で関数を作成します。

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 次のコードを追加します。 このコードについては以下の点に注目してください。

    - この `reader.readAsDataURL` メソッドは、ファイルを base64 に変換し、プロパティに格納 `reader.result` します。 メソッドが完了すると、イベント ハンドラーが `onload` トリガーされます。
    - イベント `onload` ハンドラーは、エンコードされたファイルのメタデータをトリミングし、エンコードされた文字列をグローバル変数に格納します。
    - base64 でエンコードされた文字列は、後の手順で作成した別の関数によって読み取りを行うので、グローバルに格納されます。

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

アドインは[、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)メソッドを使用してPowerPointプレゼンテーションから現在のプレゼンテーションにスライドを挿入します。 次に示すのは、ソース プレゼンテーションのすべてのスライドが現在のプレゼンテーションの先頭に挿入され、挿入されたスライドがソース ファイルの書式を保持する簡単な例です。 これは、base64 でエンコードされたバージョンのプレゼンテーション ファイルを保持する `chosenFileBase64` PowerPoint注意してください。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを 2 番目のパラメーターとして渡して、スライドが挿入される場所、ソースまたはターゲットの書式設定を取得するかどうかなど、挿入結果のいくつかの側面を `insertSlidesFromBase64` 制御できます。 次に例を示します。 このコードについては、以下の点に注意してください。

- プロパティには `formatting` 、"UseDestinationTheme" と "KeepSourceFormatting" の 2 つの値があります。 必要に応じて、列挙型 (例: ) を `InsertSlideFormatting` 使用できます `PowerPoint.InsertSlideFormatting.useDestinationTheme` 。
- この関数は、プロパティで指定されたスライドの直後に、ソース プレゼンテーションからスライドを挿入 `targetSlideId` します。 このプロパティの値は **、*nnn*#** *#* 、* mmmmmmm***、または **_nnn_ #* mmm***のいずれかの文字列で *、nnn* はスライドの ID (通常は 3 桁) で *、mmmmmmm は* スライドの作成 ID (通常は 9 桁) です。 いくつかの例は `267#763315295` 、、 `267#` 、および `#763315295` です。

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

もちろん、通常、ターゲット スライドの ID または作成 ID はコーディング時にはわかりません。 より一般的には、アドインはユーザーにターゲット スライドの選択を求める場合があります。 次の手順では、現在選択されているスライドの ***nnn*#** ID を取得し、それをターゲット スライドとして使用する方法を示します。

1. 共通 JavaScript API のOffice.context.doc[ ument.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) メソッドを使用して、現在選択されているスライドの ID を取得する関数を作成します。 次に例を示します。 呼び出しは Promise 戻り関数 `getSelectedDataAsync` に埋め込まれている点に注意してください。 これを行う理由と方法の詳細については [、「Promise-returning 関数でCommon-APIsラップ」を参照してください](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。

 
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

1. main 関数の[PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_)内で新しい関数を呼び出し、返す ID ("#" 記号と連結) をパラメーターのプロパティの値として渡 `targetSlideId` `InsertSlideOptions` します。 次に例を示します。

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

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)パラメーターを使用して、ソース プレゼンテーションから挿入されるスライドを制御することもできます。 これを行うには、ソース プレゼンテーションのスライドの ID の配列をプロパティに割り当 `sourceSlideIds` てる必要があります。 次に、4 つのスライドを挿入する例を示します。 配列内の各文字列は、プロパティに使用されるパターンの 1 つ以上に従う必要 `targetSlideId` があります。

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
> スライドは、配列に表示される順序に関係なく、ソース プレゼンテーションに表示されるのと同じ相対順序で挿入されます。

ユーザーがソース プレゼンテーションでスライドの ID または作成 ID を検出できる実用的な方法はありません。 このため、コーディング時にソースの ID を知っている場合、またはアドインが実行時に一部のデータ ソースから取得できる場合にのみ、このプロパティを `sourceSlideIds` 使用できます。 ユーザーがスライド ID を記憶できないので、ユーザーがスライド (タイトルや画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。

したがって、このプロパティは主にプレゼンテーション テンプレートのシナリオで使用されます。アドインは、挿入できるスライドのプールとして機能する特定のプレゼンテーション セットを操作するように `sourceSlideIds` 設計されています。 このようなシナリオでは、ユーザーまたは顧客のどちらかが、選択基準 (タイトルや画像など) と、可能なソース プレゼンテーションのセットから構築されたスライドの ID またはスライド作成の ID を関連付けるデータ ソースを作成および管理する必要があります。
