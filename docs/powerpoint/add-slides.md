---
title: PowerPoint でスライドを追加および削除する
description: スライドを追加および削除し、新しいスライドのマスターとレイアウトを指定する方法について説明します。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf22c18cf4089bab9091be3f4274f67974662a3
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958314"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>PowerPoint でスライドを追加および削除する

PowerPoint アドインでは、プレゼンテーションにスライドを追加し、必要に応じて、新しいスライドに使用するスライド マスターとマスターのレイアウトを指定できます。 アドインは、スライドを削除することもできます。

スライドを追加するための API は、主に、プレゼンテーション内のスライド マスターとレイアウトの ID がコーディング時にわかっているか、実行時にデータ ソースで見つかるシナリオで使用されます。 このようなシナリオでは、選択条件 (スライド マスターやレイアウトの名前や画像など) とスライド マスターとレイアウトの ID を関連付けるデータ ソースを作成し、保持する必要があります。 API は、ユーザーが既定のスライド マスターとマスターの既定のレイアウトを使用するスライドを挿入できるシナリオや、ユーザーが既存のスライドを選択して、同じスライド マスターとレイアウトを持つ新しいスライドを作成できるシナリオでも使用できます (ただし、同じ内容ではありません)。 詳細については、「 [使用するスライド マスターとレイアウトを選択する](#select-which-slide-master-and-layout-to-use) 」を参照してください。

## <a name="add-a-slide-with-slidecollectionadd"></a>SlideCollection.add を使用してスライドを追加する

[SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) メソッドを使用してスライドを追加します。 次に示すのは、プレゼンテーションの既定のスライド マスターと、そのマスターの最初のレイアウトを使用するスライドを追加する簡単な例です。 このメソッドは、常にプレゼンテーションの最後に新しいスライドを追加します。 次に例を示します。

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>使用するスライド マスターとレイアウトを選択する

[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) パラメーターを使用して、新しいスライドに使用するスライド マスターとマスター内のレイアウトを制御します。 次に例を示します。 このコードについては、以下の点に注意してください。

- オブジェクトのプロパティの一方または両方を `AddSlideOptions` 含めることができます。
- 両方のプロパティを使用する場合は、指定したレイアウトが指定されたマスターに属している必要があります。エラーがスローされます。
- プロパティが `masterId` 存在しない場合 (またはその値が空の文字列) の場合は、既定のスライド マスターが使用され、 `layoutId` そのスライド マスターのレイアウトである必要があります。
- 既定のスライド マスターは、プレゼンテーションの最後のスライドで使用されるスライド マスターです。 (プレゼンテーションに現在スライドがない場合は、既定のスライド マスターがプレゼンテーションの最初のスライド マスターになります)。
- プロパティが `layoutId` 存在しない場合 (またはその値が空の文字列) の場合は、指定された `masterId` マスターの最初のレイアウトが使用されます。
- どちらのプロパティも 3 つの形式の文字列です。***nnnnnnnnnn*#**、**#* mmmmmmmmm***、または **_nnnnnnnnnn_#* mmmmmmm****。 *ここで、nnnnnnnn* はマスターまたはレイアウトの ID (通常は 10 桁) であり、 *mmmmmmmmm* はマスターまたはレイアウトの作成 ID (通常は 6 から 10 桁) です。 いくつかの例を次に `2147483690#2908289500`示 `2147483690#`します `#2908289500`。

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

ユーザーがスライド マスターまたはレイアウトの ID または作成 ID を検出できる実用的な方法はありません。 このため、実際には、コーディング時に `AddSlideOptions` ID を知っているか、アドインが実行時にそれらを検出できる場合にのみ、パラメーターを使用できます。 ユーザーが ID を記憶することは期待できないので、ユーザーがスライド (名前や画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。

したがって、 `AddSlideOptions` このパラメーターは主に、アドインが ID が既知のスライド マスターとレイアウトの特定のセットで動作するように設計されているシナリオで使用されます。 このようなシナリオでは、選択基準 (スライド マスター、レイアウト名、画像など) と対応する ID または作成 ID を関連付けるデータ ソースを作成して管理する必要があります。

#### <a name="have-the-user-choose-a-matching-slide"></a>ユーザーに一致するスライドを選択させる

アドインを既存のスライドで使用するスライド マスターとレイアウトの同じ組み合わせを使用する必要があるシナリオで *アドインを使用* できる場合、アドインは (1) ユーザーにスライドの選択を求め、(2) スライド マスターとレイアウトの ID を読み取ることができます。 次の手順では、ID を読み取り、マスターとレイアウトが一致するスライドを追加する方法を示します。

1. 選択したスライドのインデックスを取得する関数を作成します。 次に例を示します。 このコードについては、以下の点に注意してください。

    - Common JavaScript API の [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドを使用します。
    - 呼び出しは `getSelectedDataAsync` Promise を返す関数に埋め込まれます。 これを行う理由と方法の詳細については、「 [Promise を返す関数で共通 API をラップする」を](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)参照してください。
    - `getSelectedDataAsync` は、複数のスライドを選択できるため、配列を返します。 このシナリオでは、ユーザーが 1 つしか選択していないため、コードは最初の (0 番目の) スライドを取得します。これは選択された唯一のスライドです。
    - スライドの値は `index` 、ユーザーがサムネイル ウィンドウのスライドの横に表示する 1 ベースの値です。

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

2. スライドを追加するメイン関数の [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出します。 次に例を示します。

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

## <a name="delete-slides"></a>スライドを削除する

スライドを表す [Slide](/javascript/api/powerpoint/powerpoint.slide) オブジェクトへの参照を取得し、メソッドを呼び出して、スライドを `Slide.delete` 削除します。 次に、4 番目のスライドを削除する例を示します。

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
