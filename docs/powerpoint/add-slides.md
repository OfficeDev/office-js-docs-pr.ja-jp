---
title: スライドの追加と削除を行PowerPoint
description: スライドを追加および削除し、新しいスライドのマスターとレイアウトを指定する方法について学習します。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: b14323a13332f2b1c9e26991c2446549ff78e745
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747013"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>スライドの追加と削除を行PowerPoint

新PowerPointアドインは、プレゼンテーションにスライドを追加し、必要に応じて、どのスライド マスター、およびマスター マスターのレイアウトを新しいスライドに使用を指定できます。 アドインはスライドも削除できます。

スライドを追加する API は、主に、プレゼンテーション内のスライド マスターとレイアウトの ID がコーディング時に知られているか、実行時にデータ ソースで見つかるシナリオで使用されます。 このようなシナリオでは、選択基準 (スライド マスターやレイアウトの名前やイメージなど) とスライド マスターおよびレイアウトの ID を関連付けるデータ ソースを作成および管理する必要があります。 API は、ユーザーが既定のスライド マスターとマスター の既定のレイアウトを使用するスライドを挿入できるシナリオや、ユーザーが既存のスライドを選択して、同じスライド マスターとレイアウト (ただし、同じコンテンツではない) を持つ新しいスライドを作成できるシナリオでも使用できます。 詳細 [については、「使用するスライド マスターとレイアウトの選択](#select-which-slide-master-and-layout-to-use) 」を参照してください。

## <a name="add-a-slide-with-slidecollectionadd"></a>SlideCollection.add を使用してスライドを追加する

[SlideCollection.add メソッドを使用してスライドを追加](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))します。 次に、プレゼンテーションの既定のスライド マスターを使用するスライドと、そのマスター の最初のレイアウトを追加する簡単な例を示します。 メソッドは、プレゼンテーションの最後に常に新しいスライドを追加します。 次に例を示します。

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>使用するスライド マスターとレイアウトを選択する

[AddSlideOptions パラメーターを](/javascript/api/powerpoint/powerpoint.addslideoptions)使用して、新しいスライドに使用するスライド マスターと、マスター 内で使用するレイアウトを制御します。 次に例を示します。 このコードについては、以下の点に注意してください。

- オブジェクトのプロパティのどちらかまたは両方を含 `AddSlideOptions` めることができます。
- 両方のプロパティを使用する場合は、指定したレイアウトが指定したマスターに属している必要があります。またはエラーがスローされます。
- プロパティが `masterId` 存在しない場合 ( `layoutId` または値が空の文字列である場合)、既定のスライド マスターが使用され、そのスライド マスターのレイアウトである必要があります。
- 既定のスライド マスターは、プレゼンテーションの最後のスライドで使用されるスライド マスターです。 (プレゼンテーションに現在スライドがない場合、既定のスライド マスターはプレゼンテーションの最初のスライド マスターです。
- プロパティが `layoutId` 存在しない場合 (または値が空の文字列である) 場合は、the で指定されたマスター の最初のレイアウトが `masterId` 使用されます。
- どちらのプロパティも、***nnnnnn*#**、**#* mmmmmmm***、または **_nnmm_#****のいずれかの文字列であり、 *nnnnnnnnnnnn は* マスターまたはレイアウトの ID (通常は 10 桁) で、 *mmmmmmmmm は* マスターまたはレイアウトの作成 ID (通常は 6 ~ 10 桁) です。 いくつかの例は、、 `2147483690#2908289500`、 `2147483690#`および です `#2908289500`。

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

ユーザーがスライド マスターまたはレイアウトの ID または作成 ID を検出できる実用的な方法はありません。 このため、コーディング `AddSlideOptions` 時に ID を知っている場合、またはアドインが実行時に検出できる場合にのみ、このパラメーターを使用できます。 ユーザーが ID を記憶する必要が生じないので、ユーザーがスライド (名前または画像など) を選択し、各タイトルまたは画像をスライドの ID と関連付ける方法も必要です。

したがって、この `AddSlideOptions` パラメーターは主に、アドインが特定のスライド マスターとレイアウトのセットで動作するように設計されたシナリオで使用されます。その ID は既知です。 このようなシナリオでは、ユーザーまたは顧客のどちらかが、選択基準 (スライド マスター名やレイアウト名やイメージなど) と対応する ID または作成用の ID を関連付けるデータ ソースを作成および管理する必要があります。

#### <a name="have-the-user-choose-a-matching-slide"></a>ユーザーに一致するスライドを選択する

新しいスライドで既存のスライドで使用されるスライド マスターとレイアウトの同じ組み合わせを使用するシナリオでアドインを使用できる場合は、(1) ユーザーにスライドの選択を求めるプロンプトを表示し、(2) スライド マスターとレイアウトの ID を読み取る必要があります。 次の手順では、一致するマスターとレイアウトを持つスライドを読み取り、スライドを追加する方法を示します。

1. 選択したスライドのインデックスを取得するメソッドを作成します。 次に例を示します。 このコードについては、以下の点に注意してください。

    - 共通 JavaScript [API Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドを使用します。
    - 呼び出しは `getSelectedDataAsync` Promise 戻り関数に埋め込まれている。 これを行う理由と方法の詳細については、「Promise-returning 関数で一般的な API をラップする [」を参照してください](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。
    - `getSelectedDataAsync` 複数のスライドを選択できるので、配列を返します。 このシナリオでは、ユーザーが選択したスライドは 1 つだけなので、コードは最初の (0 番目) スライドを取得します。これが選択された唯一のスライドです。
    - スライド `index` の値は、ユーザーがサムネイル ウィンドウのスライドの横に表示する 1 ベースの値です。

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

2. スライドを追加するメイン関数[の PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) 内で新しい関数を呼び出します。 次に例を示します。

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

スライドを表す [Slide](/javascript/api/powerpoint/powerpoint.slide) オブジェクトへの参照を取得してスライドを削除し、メソッドを呼び出 `Slide.delete` します。 次に、4 番目のスライドを削除する例を示します。

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
