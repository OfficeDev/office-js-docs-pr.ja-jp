---
title: DOM とランタイム環境を読み込む
description: DOM と Office アドインのランタイム環境を読み込む
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 557297fc9e13ebab5b4eebd7917d0e0d9e444e88
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608126"
---
# <a name="loading-the-dom-and-runtime-environment"></a>DOM とランタイム環境を読み込む

アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。

## <a name="startup-of-a-content-or-task-pane-add-in"></a>コンテンツまたは作業ウィンドウ アドインの起動

次の図では、Excel、PowerPoint、Project、または Word のコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。

![コンテンツ アドインまたは作業ウィンドウ アドイン起動時のイベントのフロー](../images/office15-app-sdk-loading-dom-agave-runtime.png)

コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。

1. ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。

2. Office ホスト アプリケーションが、アドインの XML マニフェストを AppSource、SharePoint のアプリ カタログ、またはアドインの発生元である共有フォルダー カタログから読み取ります。

3. Office ホスト アプリケーションが、ブラウザー コントロールにアドインの HTML ページを開きます。

    次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。

4. ブラウザーコントロールが DOM と HTML 本文を読み込み、イベントのイベントハンドラーを呼び出し `window.onload` ます。

5. Office ホスト アプリケーションがランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、[Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するアドインのイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。 との違いの詳細については `Office.initialize` `Office.onReady` 、「[アドインを初期化する](initialize-add-in.md)」を参照してください。

6. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。


## <a name="startup-of-an-outlook-add-in"></a>Outlook アドインの起動

次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。

![Outlook アドイン起動時のイベントのフロー](../images/outlook15-loading-dom-agave-runtime.png)

Outlook アドインが起動すると、次のイベントが発生します。

1. Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。

2. ユーザーが Outlook でアイテムを選択します。

3. 選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。

4. ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。

5. ブラウザーコントロールが DOM と HTML 本文を読み込み、イベントのイベントハンドラーを呼び出し `onload` ます。

6. Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。 との違いの詳細については `Office.initialize` `Office.onReady` 、「[アドインを初期化する](initialize-add-in.md)」を参照してください。

7. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。


## <a name="checking-the-load-status"></a>読み込み状態のチェック

DOM とランタイム環境の両方で読み込みが完了したことを確認する方法の 1 つは、jQuery [.ready()](https://api.jquery.com/ready/) 関数を使用することです: `$(document).ready()`。 たとえば、次の `onReady` イベントハンドラーは、アドインの初期化に固有のコードが実行される前に、DOM が最初に読み込まれることを確認します。 その後、 `onReady` ハンドラーは[メールボックス. item](/javascript/api/outlook/office.mailbox#item)プロパティを使用して、Outlook で現在選択されているアイテムを取得し、アドインの main 関数を呼び出し `initDialer` ます。

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

または、 `initialize` 次の例に示すように、同じコードをイベントハンドラーで使用することもできます。

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

この方法は、 `onReady` `initialize` Office アドインのハンドラーでも使用できます。

ダイヤラー サンプル Outlook アドインでは、JavaScript のみを使用してこれらと同じ条件を確認するという少し異なる方法を使用しています。

> [!IMPORTANT]
> アドインに実行する初期化タスクがない場合でも、 `Office.onReady` 次の例に示されているように、少なくとも最小のイベントハンドラー関数の呼び出しを含める必要があり `Office.initialize` ます。
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> `Office.onReady`イベントハンドラーの呼び出しまたは割り当てを行わない場合 `Office.initialize` 、アドインの起動時にエラーが発生することがあります。 また、ユーザーが Excel、PowerPoint、または Outlook などの Office Web クライアントでアドインを使用しようとすると、実行に失敗します。
>
> アドインに複数のページが含まれている場合は、新しいページが読み込まれるときに、そのページがイベントハンドラーを呼び出すか、または割り当てる必要があり `Office.onReady` `Office.initialize` ます。

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office アドインを初期化する](initialize-add-in.md)
