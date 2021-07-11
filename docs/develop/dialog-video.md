---
title: Office ダイアログ ボックスを使用してビデオを再生する
description: '[ビデオの再生] ダイアログ ボックスでビデオを開いて再生するOffice説明します。'
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 2519b2f105503a0479eee07d885a1543f5455343
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349883"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>[ビデオをOffice]ダイアログ ボックスを使用してビデオを表示する

この記事では、[アドイン] ダイアログ ボックスでビデオを再生Office説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って[、Office](dialog-api-in-office-add-ins.md)ダイアログ ボックスを使用する基本について理解している必要があります。

ダイアログ API を使用してダイアログ ボックスでビデオを再生するにはOffice手順を実行します。

1. iframe と他のコンテンツを含むページを作成します。 ページはホスト ページと同じドメインにある必要があります。 ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。 `src`iframe の属性で、オンライン ビデオの URL をポイントします。 ビデオの URL のプロトコルは HTTPS である必要があります。 この記事では、このページを "l" と呼video.dialogbox.htmします。 マークアップの例を次に示します。

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
