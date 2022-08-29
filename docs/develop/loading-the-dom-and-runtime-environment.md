---
title: DOM とランタイム環境を読み込む
description: DOM と Office アドインのランタイム環境を読み込みます。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 707b6f6f743767571cf0ab7f465ddf84f117a63b
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423077"
---
# <a name="load-the-dom-and-runtime-environment"></a>DOM とランタイム環境を読み込む

独自のカスタム ロジックを実行する前に、アドインは DOM と Office アドインの [両方のランタイム](../testing/runtimes.md) 環境を確実に読み込む必要があります。

## <a name="startup-of-a-content-or-task-pane-add-in"></a>コンテンツまたは作業ウィンドウ アドインの起動

次の図では、Excel、PowerPoint、Project、または Word のコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。

![コンテンツまたは作業ウィンドウ アドインを開始するときのイベントのフロー。](../images/office15-app-sdk-loading-dom-agave-runtime.png)

次のイベントは、コンテンツまたは作業ウィンドウアドインが起動したときに発生します。

1. ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。

2. Office クライアント アプリケーションは、AppSource、SharePoint 上のアプリ カタログ、またはそれが作成された共有フォルダー カタログからアドインの XML マニフェストを読み取ります。

3. Office クライアント アプリケーションは、ブラウザー コントロールでアドインの HTML ページを開きます。

    次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。

4. ブラウザー コントロールは DOM と HTML の本文を読み込み、イベントのイベント ハンドラーを `window.onload` 呼び出します。

5. Office クライアント アプリケーションは、コンテンツ配布ネットワーク (CDN) サーバーから Office JavaScript API ライブラリ ファイルをダウンロードしてキャッシュするランタイム環境を読み込み、ハンドラーが割り当てられている場合は、[Office](/javascript/api/office) オブジェクトの[初期化](/javascript/api/office#Office_initialize_reason_)イベントに対してアドインのイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンメソッド `then()` ) がハンドラーに渡された (またはチェーンされた) `Office.onReady` かどうかを確認します。 アドインの`Office.onReady`違`Office.initialize`いの詳細については、「[アドインを初期化](initialize-add-in.md)する」を参照してください。

6. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。

## <a name="startup-of-an-outlook-add-in"></a>Outlook アドインの起動

次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。

![Outlook アドインを起動するときのイベントのフロー。](../images/outlook15-loading-dom-agave-runtime.png)

次のイベントは、Outlook アドインが起動したときに発生します。

1. Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。

2. ユーザーが Outlook でアイテムを選択します。

3. 選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。

4. ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。

5. ブラウザー コントロールは DOM と HTML の本文を読み込み、イベントのイベント ハンドラーを `onload` 呼び出します。

6. Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#Office_initialize_reason_) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。 現時点では、コールバック (またはチェーンメソッド `then()` ) がハンドラーに渡された (またはチェーンされた) `Office.onReady` かどうかを確認します。 アドインの`Office.onReady`違`Office.initialize`いの詳細については、「[アドインを初期化](initialize-add-in.md)する」を参照してください。

7. DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office アドインを初期化する](initialize-add-in.md)
- [Office アドインのランタイム](../testing/runtimes.md)
