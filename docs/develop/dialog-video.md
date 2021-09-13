---
title: Office ダイアログ ボックスを使用してビデオを再生する
description: '[ビデオの再生] ダイアログ ボックスでビデオを開いて再生するOffice説明します。'
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4765c9d549bfcc52315420342bf99ec88fdebf24
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151441"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>[ビデオをOffice]ダイアログ ボックスを使用してビデオを表示する

この記事では、[アドイン] ダイアログ ボックスでビデオを再生Office説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って[、Office](dialog-api-in-office-add-ins.md)ダイアログ ボックスを使用する基本について理解している必要があります。

ダイアログ API を使用してダイアログ ボックスでビデオを再生するにはOffice手順を実行します。

1. iframe と他のコンテンツを含むページを作成します。 ページはホスト ページと同じドメインにある必要があります。 ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。 `src`iframe の属性で、オンライン ビデオの URL をポイントします。 ビデオの URL のプロトコルは HTTPS である必要があります。 この記事では、このページを "video.dialogbox.html" と呼video.dialogbox.html。 マークアップの例を次に示します。

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. ホスト ページで `displayDialogAsync` の呼び出しを使用して、video.dialogbox.html を開きます。
3. ユーザーがダイアログ ボックスを閉じたときに、アドインに通知する必要がある場合は、`DialogEventReceived` イベントのハンドラーを登録して、12006 イベントを処理します。 詳細については、「エラーと[イベント」ダイアログ ボックスOffice参照してください](dialog-handle-errors-events.md)。

ダイアログ ボックスで再生するビデオのサンプルについては、ビデオの配置パターン [を参照してください](../design/first-run-experience-patterns.md#video-placemat)。

![アプリの前にあるアドイン ダイアログ ボックスで再生されているビデオを示すExcel。](../images/video-placemats-dialog-open.png)
