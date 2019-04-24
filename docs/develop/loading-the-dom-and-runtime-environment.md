---
title: DOM とランタイム環境を読み込む
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b1f63d9fe012ed8c8a5cf4a0f7de862ddabcd4d3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449845"
---
# <a name="loading-the-dom-and-runtime-environment"></a>DOM とランタイム環境を読み込む

アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。 

## <a name="startup-of-a-content-or-task-pane-add-in"></a>コンテンツまたは作業ウィンドウ アドインの起動

次の図は、Excel、PowerPoint、Project、Word、または Access でコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。

![コンテンツ アドインまたは作業ウィンドウ アドイン起動時のイベントのフロー](../images/office15-app-sdk-loading-dom-agave-runtime.png)

コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。

1. ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。

2. Office ホスト アプリケーションが、アドインの XML マニフェストを AppSource、SharePoint のアドイン カタログ、またはアドインの発生元である共有フォルダー カタログから読み取ります。

3. Office ホスト アプリケーションが、ブラウザー コントロールにアドインの HTML ページを開きます。

    次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。

4. ブラウザー コントロールが、DOM と HTML 本文を読み込み、**window.onload** イベントに対するイベント ハンドラーを呼び出します。

5. Office ホスト アプリケーションがランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、[Office](/javascript/api/office#office-initialize) オブジェクトの [initialize](/javascript/api/office) イベントに対するアドインのイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。 `Office.initialize` と `Office.onReady` の違いの詳細については、「[アドインの初期化](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)」をご覧ください。

6. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。


## <a name="startup-of-an-outlook-add-in"></a>Outlook アドインの起動

次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。

![Outlook アドイン起動時のイベントのフロー](../images/outlook15-loading-dom-agave-runtime.png)

Outlook アドインが起動すると、次のイベントが発生します。

1. Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。

2. ユーザーが Outlook でアイテムを選択します。

3. 選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。

4. ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。

5. ブラウザー コントロールが DOM と HTML 本文を読み込んで、**onload** イベントに対するイベント ハンドラーを呼び出します。

6. Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。 `Office.initialize` と `Office.onReady` の違いの詳細については、「[アドインの初期化](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)」をご覧ください。

7. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。


## <a name="checking-the-load-status"></a>読み込み状態のチェック

DOM とランタイム環境の両方で読み込みが完了したことを確認する方法の 1 つは、jQuery [.ready()](https://api.jquery.com/ready/) 関数を使用することです: `$(document).ready()`。 たとえば、次の **onReady** イベント ハンドラーは、アドインの実行の初期化に固有のコードの前に、DOM が最初に読み込まれることを確認します。 その後、**onReady** ハンドラーは [mailbox.item](/javascript/api/outlook/office.mailbox) プロパティを使用して、Outlook で現在選択されている項目を取得し、アドインのメイン関数 `initDialer` を呼び出します。

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

また、次の例に示されているように、同じコードを **initialize** イベント ハンドラーで使用することができます。

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

この方法は、任意の Office アドインの **onReady** または **initialize** ハンドラーで使用できます。

ダイヤラー サンプル Outlook アドインでは、JavaScript のみを使用してこれらと同じ条件を確認するという少し異なる方法を使用しています。 

> [!IMPORTANT]
> アドインに実行する初期化タスクがない場合でも、次の例に示されているように、**Office.onReady** の呼び出しを少なくとも 1 つ含めるか、最小のイベント ハンドラー関数 **Office.initialize** を割り当てる必要があります。
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> **Office.onReady** を呼び出したり、**Office.initialize** イベントを割り当てたりしない場合、アドインを開始するとエラーが発生する可能性があります。 また、ユーザーが Excel Online、PowerPoint Online、Outlook Web App などの Office Online Web クライアントでアドインを使用しようとすると、アドインの実行が失敗します。
>
> アドインに複数のページが含まれる場合、新しいページが読み込まれるときに、そのページは **Office.onReady** を呼び出すか、**Office.initialize** イベント ハンドラーを割り当てる必要があります。

## <a name="see-also"></a>関連項目

- [JavaScript API for Office について](understanding-the-javascript-api-for-office.md)
