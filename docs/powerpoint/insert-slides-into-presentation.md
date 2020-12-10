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
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a>PowerPoint プレゼンテーションでスライドを挿入および削除する (プレビュー)

PowerPoint アドインでは、PowerPoint のアプリケーション固有の JavaScript ライブラリを使用して、プレゼンテーションのスライドを現在のプレゼンテーションに挿入できます。 挿入したスライドに、元のプレゼンテーションの書式設定を保持するか、または対象となるプレゼンテーションの書式設定を保持するかを制御できます。 プレゼンテーションからスライドを削除することもできます。

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

スライド挿入 Api は、主にプレゼンテーションテンプレートのシナリオで使用されます。これは、アドインによって挿入できるスライドのプールとして機能する既知のプレゼンテーションの数が少ないことです。 このようなシナリオでは、ユーザーまたは顧客は、選択基準 (スライドタイトル、画像など) とスライド Id を関連付けたデータソースを作成して維持する必要があります。 この Api は、ユーザーが任意のプレゼンテーションからスライドを挿入できる場合にも使用できますが、このシナリオでは、ユーザーは、元のプレゼンテーションの *すべて* のスライドを挿入することに効果的に制限されます。 詳細については、「 [挿入するスライドを選択](#selecting-which-slides-to-insert) する」を参照してください。

プレゼンテーションのスライドを別のプレゼンテーションに挿入するには、2つの手順を実行します。

1. ソースプレゼンテーションファイル (.pptx) を base64 形式の文字列に変換します。
1. メソッドを使用して、 `insertSlidesFromBase64` base64 ファイルから現在のプレゼンテーションに1つまたは複数のスライドを挿入します。

## <a name="convert-the-source-presentation-to-base64"></a>ソースプレゼンテーションを base64 に変換する

ファイルを base64 に変換するには、さまざまな方法があります。 使用するプログラミング言語とライブラリ、およびアドインまたはクライアント側のサーバー側で変換するかどうかは、シナリオによって決まります。 通常、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) オブジェクトを使用して、クライアント側の JavaScript で変換を行います。 次の例は、この方法を示しています。

1. 最初に、ソースの PowerPoint ファイルへの参照を取得します。 この例では、 `<input>` 種類のコントロールを使用し `file` て、ユーザーにファイルの選択を求めるメッセージを表示します。 次のマークアップをアドインページに追加します。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    このマークアップは、次のスクリーンショットの UI をページに追加します。

    ![HTML ファイルの種類の入力コントロールの前に説明文を表示するスクリーンショット。「スライドを挿入する PowerPoint プレゼンテーションを選択する」を参照してください。 このコントロールは、"Choose file" というラベルの付いたボタンで構成され、その後に "ファイルが選択されていません" という文が続きます。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > PowerPoint ファイルを取得する方法は他にもたくさんあります。 たとえば、ファイルが OneDrive または SharePoint に保存されている場合は、Microsoft Graph を使用してダウンロードできます。 詳細については、「 [Microsoft graph でファイルを処理](/graph/api/resources/onedrive) する」および「 [Microsoft graph でファイルにアクセス](/learn/modules/msgraph-access-file-data/)する」を参照してください。

2. 次のコードをアドインの JavaScript に追加して、入力コントロールのイベントに関数を割り当て `change` ます。 (この関数は、 `storeFileAsBase64` 次の手順で作成します)。

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 次のコードを追加します。 このコードについては、次の点に注意してください。

    - `reader.readAsDataURL`このメソッドは、ファイルを base64 に変換して、プロパティに格納し `reader.result` ます。 メソッドが完了すると、イベントハンドラーがトリガーされ `onload` ます。
    - イベントハンドラーは、エンコードされた `onload` ファイルからメタデータをトリミングし、エンコードされた文字列をグローバル変数に格納します。
    - Base64 でエンコードされた文字列は、後の手順で作成した別の関数によって読み取られるため、グローバルに格納されます。

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>InsertSlidesFromBase64 を使用してスライドを挿入する

アドインは、 [insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) メソッドを使用して、別の PowerPoint プレゼンテーションのスライドを現在のプレゼンテーションに挿入します。 次の簡単な例は、元のプレゼンテーションのすべてのスライドを現在のプレゼンテーションの先頭に挿入し、挿入したスライドにソースファイルの書式を保持します。 これ `chosenFileBase64` は、PowerPoint プレゼンテーションファイルの base64 でエンコードされたバージョンを保持するグローバル変数であることに注意してください。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)オブジェクトを2番目のパラメーターとして渡すことによって、挿入結果のいくつかの側面を制御できます。これには、スライドが挿入される場所や、ソースまたはターゲットの書式設定を取得するかどうかなどがあり `insertSlidesFromBase64` ます。 次に例を示します。 このコードについては、以下の点に注意してください。

- プロパティには、 `formatting` "UseDestinationTheme" と "KeepSourceFormatting" という2つの値があります。 必要に応じて、 `InsertSlideFormatting` enum (など) を使用でき `PowerPoint.InsertSlideFormatting.useDestinationTheme` ます。
- 関数は、プロパティで指定されたスライドの直後に、元のプレゼンテーションのスライドを挿入し `targetSlideId` ます。 このプロパティの値は、次の3つの形式のうちの1つです。 ***nnn * #**、* *#* mmmmmmmmm * * *、または **_nnn_ #* mmmmmmmmm * * *、 *nnn* はスライドの id (通常は3桁)、 *mmmmmmmmm* はスライドの作成 ID (通常は9桁) です。 例として、、、などがあり `267#763315295` `267#` `#763315295` ます。

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

当然のことですが、通常は、ターゲットスライドの ID または作成 ID のコーディング時にはわかりません。 一般的に、アドインは、ターゲットスライドを選択するようにユーザーに要求します。 次の手順は、現在選択されているスライドの ***nnn * #** ID を取得し、それをターゲットスライドとして使用する方法を示しています。

1. 共通 JavaScript Api の [Office.context.document](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) を使用して、現在選択されているスライドの ID を取得する関数を作成します。 次に例を示します。 の呼び出し `getSelectedDataAsync` は、Promise を返す関数に埋め込まれていることに注意してください。 その理由と方法の詳細については、「 [関数を返す関数のラップ Common-APIs](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)」を参照してください。

 
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

1. Main 関数の [PowerPoint. run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出し、返される ID ("#" 記号で連結される) を `targetSlideId` パラメーターのプロパティの値として渡します `InsertSlideOptions` 。 次に例を示します。

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

[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)パラメーターを使用して、ソースプレゼンテーションから挿入するスライドを制御することもできます。 これを行うには、移動元のプレゼンテーションのスライド Id の配列を `sourceSlideIds` プロパティに割り当てます。 次に、4つのスライドを挿入する例を示します。 配列内の各文字列は、プロパティに使用されている1つまたは複数のパターンに従う必要があることに注意してください `targetSlideId` 。

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
> スライドは、配列に表示される順序に関係なく、ソースプレゼンテーションに表示される相対的な順序で挿入されます。

ユーザーがソースプレゼンテーションのスライドの ID または作成 ID を検出するための実用的な方法はありません。 このため、このプロパティは、 `sourceSlideIds` コーディング時にソース id がわかっている場合、またはデータソースから実行時にアドインで取得できる場合にのみ使用できます。 ユーザーがスライド Id を記憶することは予想できないため、ユーザーがスライドを選択できるようにするには、タイトルや画像によるスライドの選択が可能になり、各タイトルまたは画像をスライドの ID と関連付けることが必要になります。

そのため、この `sourceSlideIds` プロパティは、主にプレゼンテーションテンプレートのシナリオで使用されます。アドインは、挿入可能なスライドのプールとして機能する特定のプレゼンテーションセットで機能するように設計されています。 このようなシナリオでは、お客様またはお客様は、選択基準 (タイトル、画像など) を含むデータソースを作成して、使用可能なソースプレゼンテーションセットから構成されているスライド作成 Id を関連付けて保持する必要があります。

## <a name="delete-slides"></a>スライドの削除

スライドを表す [slide オブジェクトへの参照](/javascript/api/powerpoint/powerpoint.slide) を取得して、メソッドを呼び出すことで、スライドを削除でき `Slide.delete` ます。 次に、4番目のスライドを削除する例を示します。

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
