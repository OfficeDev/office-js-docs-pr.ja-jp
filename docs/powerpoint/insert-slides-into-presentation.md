---
title: PowerPoint プレゼンテーションにスライドを挿入する
description: プレゼンテーションのスライドを別のプレゼンテーションに挿入する方法について説明します。
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: a31933de4272634394dc6c36aafa973c41265471
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857572"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>PowerPoint プレゼンテーションにスライドを挿入する

PowerPoint アドインは、PowerPoint のアプリケーション固有の JavaScript ライブラリを使用して、1 つのプレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。 挿入されたスライドがソース プレゼンテーションの書式設定を保持するか、ターゲット プレゼンテーションの書式設定を保持するかを制御できます。

スライド挿入 API は、主にプレゼンテーション テンプレートのシナリオで使用されます。アドインで挿入できるスライドのプールとして機能する既知のプレゼンテーションが少数あります。 このようなシナリオでは、選択基準 (スライド タイトルや画像など) とスライド ID を関連付けるデータ ソースを作成して管理する必要があります。 API は、ユーザーが任意のプレゼンテーションからスライドを挿入できるシナリオでも使用できますが、そのシナリオでは、ユーザーは実質的にソース プレゼンテーション *からすべての* スライドを挿入することに制限されます。 詳細については、「 [挿入するスライドの選択](#selecting-which-slides-to-insert) 」を参照してください。

あるプレゼンテーションから別のプレゼンテーションにスライドを挿入するには、2 つの手順があります。

1. ソース プレゼンテーション ファイル (.pptx) を base64 形式の文字列に変換します。
1. base64 ファイルから現在の `insertSlidesFromBase64` プレゼンテーションに 1 つ以上のスライドを挿入するには、このメソッドを使用します。

## <a name="convert-the-source-presentation-to-base64"></a>ソース プレゼンテーションを base64 に変換する

ファイルを base64 に変換する方法は多数あります。 使用するプログラミング言語とライブラリ、およびアドインのサーバー側またはクライアント側で変換するかどうかは、シナリオによって決まります。 最も一般的には、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側の JavaScript で変換を行います。 次の例は、この方法を示しています。

1. まず、ソース PowerPoint ファイルへの参照を取得します。 この例では、種類`file`のコントロールを`<input>`使用して、ユーザーにファイルの選択を求めます。 アドイン ページに次のマークアップを追加します。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    このマークアップは、次のスクリーンショットの UI をページに追加します。

    ![HTML ファイルの種類の入力コントロールの前に、「スライドを挿入する PowerPoint プレゼンテーションを選択する」という説明文が表示されているスクリーンショット。 コントロールは、"ファイルの選択" というラベルの付いたボタンと、"ファイルが選択されていません" という文で構成されます。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > PowerPoint ファイルを取得するには、他にも多くの方法があります。 たとえば、ファイルが OneDrive または SharePoint に保存されている場合は、Microsoft Graph を使用してダウンロードできます。 詳細については、「[Microsoft Graph でのファイルの操作」と「Microsoft Graph を使用](/graph/api/resources/onedrive)[したファイルへのアクセス](/training/modules/msgraph-access-file-data/)」を参照してください。

2. アドインの JavaScript に次のコードを追加して、入力コントロールの `change` イベントに関数を割り当てます。 (次の `storeFileAsBase64` 手順で関数を作成します)。

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 次のコードを追加します。 このコードについては、次の点に注意してください。

    - このメソッドは `reader.readAsDataURL` 、ファイルを base64 に変換し、プロパティに `reader.result` 格納します。 メソッドが完了すると、イベント ハンドラーが `onload` トリガーされます。
    - イベント ハンドラーは `onload` 、エンコードされたファイルからメタデータをトリミングし、エンコードされた文字列をグローバル変数に格納します。
    - base64 でエンコードされた文字列は、後の手順で作成した別の関数によって読み取られるので、グローバルに格納されます。

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>insertSlidesFromBase64 でスライドを挿入する

アドインは、 [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) メソッドを使用して、別の PowerPoint プレゼンテーションのスライドを現在のプレゼンテーションに挿入します。 次の簡単な例では、ソース プレゼンテーションのすべてのスライドが現在のプレゼンテーションの先頭に挿入され、挿入されたスライドはソース ファイルの書式設定を保持します。 PowerPoint `chosenFileBase64` プレゼンテーション ファイルの base64 でエンコードされたバージョンを保持するグローバル変数であることに注意してください。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) オブジェクトを 2 番目のパラメーターとして渡すことで、スライドが挿入される場所やソースまたはターゲットの書式設定を取得するかどうかなど、挿入結果のいくつかの側面を`insertSlidesFromBase64`制御できます。 次に例を示します。 このコードについては、以下の点に注意してください。

- プロパティには、"UseDestinationTheme" と "KeepSourceFormatting" の 2 つの値 `formatting` があります。 必要に応じて、列挙型 (例: `PowerPoint.InsertSlideFormatting.useDestinationTheme`) を使用`InsertSlideFormatting`できます。
- この関数は、プロパティで指定されたスライドの直後に、ソース プレゼンテーションからスライドを `targetSlideId` 挿入します。 このプロパティの値は、可能な 3 つの形式のいずれかです。***nnn*#**、**#* mmmmmmmmm***、または *nnn mmmmmmmmm***。*_nnn_#* はスライドの ID (通常は 3 桁) で *、mmmmmmmmm* はスライドの作成 ID (通常は 9 桁) です。 いくつかの例を次に `267#763315295`示 `267#`します `#763315295`。

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

もちろん、通常、ターゲット スライドの ID または作成 ID はコーディング時にはわかりません。 より一般的には、アドインはユーザーにターゲット スライドの選択を求めます。 次の手順では、現在選択されているスライドの ***nnn*#** ID を取得し、それをターゲット スライドとして使用する方法を示します。

1. Common JavaScript API の [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドを使用して、現在選択されているスライドの ID を取得する関数を作成します。 次に例を示します。 呼び出しの `getSelectedDataAsync` 対象は Promise を返す関数に埋め込まれている点に注意してください。 これを行う理由と方法の詳細については、「 [promise-returning 関数でCommon-APIsをラップする](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)」を参照してください。

 
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

1. メイン関数の [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出し、返される ID をパラメーターのプロパティ`InsertSlideOptions`の値`targetSlideId`として渡します ("#" 記号で連結されます)。 次に例を示します。

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

### <a name="selecting-which-slides-to-insert"></a>挿入するスライドを選択する

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) パラメーターを使用して、ソース プレゼンテーションからどのスライドを挿入するかを制御することもできます。 これを行うには、ソース プレゼンテーションのスライド ID の配列をプロパティに `sourceSlideIds` 割り当てます。 次に、4 つのスライドを挿入する例を示します。 配列内の各文字列は、プロパティに使用される `targetSlideId` パターンの 1 つまたは複数に従う必要があることに注意してください。

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

ユーザーがソース プレゼンテーションでスライドの ID または作成 ID を検出できる実用的な方法はありません。 このため、このプロパティは、コーディング時にソース ID がわかっているか、アドインが実行時にデータ ソースから取得できる場合にのみ使用 `sourceSlideIds` できます。 ユーザーがスライド ID を記憶することは期待できないため、ユーザーがスライド (タイトルや画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。

したがって、 `sourceSlideIds` このプロパティは主にプレゼンテーション テンプレートのシナリオで使用されます。アドインは、挿入できるスライドのプールとして機能する特定のプレゼンテーション セットを操作するように設計されています。 このようなシナリオでは、ユーザーまたは顧客が、選択基準 (タイトルや画像など) と、使用可能なソース プレゼンテーションのセットから作成されたスライド ID またはスライド作成 ID を関連付けるデータ ソースを作成および管理する必要があります。
