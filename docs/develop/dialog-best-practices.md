---
title: Office ダイアログ API のベスト プラクティスとルール
description: 単一ページ アプリケーション (SPA) Officeベスト プラクティスなど、新しいダイアログ API のルールとベスト プラクティスを提供します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 4359d116e9720255278c5b3f543b135013c7e76c
ms.sourcegitcommit: 7cd501d0fdbbd4636bd08647b638dd5ca4c7c630
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/17/2021
ms.locfileid: "50282983"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office ダイアログ API のベスト プラクティスとルール

この記事では、ダイアログの UI を設計し、単一ページ アプリケーション (SPA) で API を使用する場合のベスト プラクティスなど、Office ダイアログ API のルール、説明、ベスト プラクティスについて説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、Office ダイアログ [API](dialog-api-in-office-add-ins.md)の使用の基本を理解している必要があります。
> 
> また、[ [エラーとイベントの処理] ダイアログ ボックスOffice参照してください](dialog-handle-errors-events.md)。

## <a name="rules-and-gotchas"></a>ルールと注意事項

- ダイアログ ボックスは、HTTP ではなく HTTPS URL にのみ移動できます。
- [displayDialogAsync](/javascript/api/office/office.ui)メソッドに渡される URL は、アドイン自体とまったく同じドメインにある必要があります。 サブドメインにすることはできません。 ただし、このページに渡されたページは、別のドメイン内のページにリダイレクトできます。
- ホスト ウィンドウは、作業ウィンドウでも、アドイン コマンドの UI[](../reference/manifest/functionfile.md)なし関数ファイルでも使用できます。一度に開くことができるダイアログ ボックスは 1 つのみです。
- ダイアログ ボックスOffice呼び出し可能な API は 2 つのみです。
  - [messageParent](/javascript/api/office/office.ui#messageparent-message-)関数。
  - `Office.context.requirements.isSetSupported` (詳細については、「アプリケーションと [API の要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md))。
- [messageParent](/javascript/api/office/office.ui#messageparent-message-)関数は、アドイン自体とまったく同じドメイン内のページからのみ呼び出しできます。

## <a name="best-practices"></a>ベスト プラクティス

### <a name="avoid-overusing-dialog-boxes"></a>ダイアログ ボックスの使い過ぎを避ける

UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については [、Excel アドインの JavaScript SalesTracker サンプルを参照](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) してください。

### <a name="designing-a-dialog-box-ui"></a>ダイアログ ボックス UI の設計

ダイアログ ボックスの設計のベスト プラクティスについては、アドインのダイアログ ボックス [Office参照してください](../design/dialog-boxes.md)。

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Office on the web を使用したポップアップ ブロックの処理

web 上で Office を使用している間にダイアログ ボックスを表示しようとすると、ブラウザーのポップアップ ブロックがダイアログ ボックスをブロックする可能性があります。 Officeには、アドインのダイアログ ボックスをブラウザーのポップアップ ブロックの例外にする機能があります。 コードがメソッドを呼び出す場合、web Officeをクリックすると、次のような `displayDialogAsync` プロンプトが開きます。

![ブラウザー内のポップアップ ブロックを回避するためにアドインが生成できる簡単な説明と [許可] ボタンと [無視] ボタンを示すプロンプトを示すスクリーンショット](../images/dialog-prompt-before-open.png)

ユーザーが [許可] を **選択すると**、[Officeが開きます。 ユーザーが [無視 **]** を選択すると、プロンプトが閉じOfficeダイアログ ボックスが開かれません。 代わりに、この `displayDialogAsync` メソッドはエラー 12009 を返します。 コードは、このエラーをキャッチして、ダイアログを必要としない代替エクスペリエンスを提供するか、アドインがダイアログを許可する必要があるというメッセージをユーザーに表示する必要があります。 (12009 の詳細については [、「displayDialogAsync からのエラー」を参照](dialog-handle-errors-events.md#errors-from-displaydialogasync)してください)。

何らかの理由でこの機能をオフにする場合は、コードをオプトアウトする必要があります。この要求は、メソッドに渡 [される DialogOptions](/javascript/api/office/office.dialogoptions) オブジェクトで行 `displayDialogAsync` います。 具体的には、オブジェクトに含める必要があります `promptBeforeOpen: false` 。 このオプションが false に設定されている場合、Office on the web はアドインがダイアログを開くことを許可するように求めるメッセージをユーザーに表示し、Office ダイアログは開かれません。

### <a name="do-not-use-the-_host_info-value"></a>ホスト情報の値 \_ を \_ 使用しない

Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 カスタム クエリ パラメーターがある場合は、その後に追加されます。 ダイアログ ボックスが移動する後続の URL には追加されません。 Microsoft は、この値の内容を変更したり、完全に削除したりする可能性があります。そのため、コードで読み取る必要はありません。 ダイアログ ボックスのセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に同じ値が追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

### <a name="opening-another-dialog-immediately-after-closing-one"></a>1 つを閉じる直後に別のダイアログを開く

特定のホスト ページから複数のダイアログを開く方法はないので、別のダイアログを開く前に、開いているダイアログで [Dialog.close](/javascript/api/office/office.dialog#close__) を呼び出 `displayDialogAsync` す必要があります。 メソッド `close` は非同期です。 このため、呼び出しの直後に呼び出しを行った場合、2 番目のダイアログを開Office最初のダイアログが完全に閉じてい `displayDialogAsync` `close` なOffice可能性があります。 この場合、Office [12007](dialog-handle-errors-events.md#12007) エラーが返されます。"このアドインには既にアクティブなダイアログが含まれるため、操作は失敗しました。"

メソッドはコールバック パラメーターを受け入れないので、Promise オブジェクトを返すので、キーワードまたはメソッドを使用して待 `close` `await` つ `then` 必要はありません。 このため、ダイアログを閉じる直後に新しいダイアログを開く必要がある場合は、メソッドで新しいダイアログを開くコードをカプセル化し、戻り値の呼び出し時にメソッドを再帰的に呼び出すメソッドを設計する方法を推奨します `displayDialogAsync` `12007` 。 次に例を示します。

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

または [、setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) メソッドを使用して 2 番目のダイアログを開く前に、コードを強制的に一時停止できます。 次に例を示します。

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>SPA で新しいOffice API を使用するためのベスト プラクティス

単一ページ アプリケーション (SPA) が通常行うので、アドインがクライアント側ルーティングを使用する場合は、ルートの URL を別の HTML ページの URL ではなく [displayDialogAsync](/javascript/api/office/office.ui) メソッドに渡すオプションがあります。 *以下の理由により、このような方法はお勧めしません。*

> [!NOTE]
> この記事は、Express *ベース* の Web アプリケーションなど、サーバー側のルーティングには関係ありません。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SPA と Office ダイアログ API に関する問題

このOfficeダイアログ ボックスは、JavaScript エンジンの独自のインスタンスを持つ新しいウィンドウ内にあるため、独自の完全な実行コンテキストです。 ルートを渡した場合、基本ページとその初期化とブートストラップ コードはすべて、この新しいコンテキストで再び実行され、すべての変数がダイアログ ボックスの初期値に設定されます。 したがって、この手法では、アプリケーションの 2 番目のインスタンスをダウンロードしてボックス ウィンドウで起動します。これにより、SPA の目的が部分的に取り上げらなされます。 さらに、ダイアログ ボックス ウィンドウで変数を変更するコードでは、同じ変数の作業ウィンドウのバージョンは変更されません。 同様に、ダイアログ ボックス ウィンドウには、作業ウィンドウのコードからアクセスできない独自のセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) があります。 ダイアログ ボックスと、呼び出されたホスト ページ `displayDialogAsync` は、サーバーに対する 2 つの異なるクライアントのように表示されます。 (ホスト ページの説明については、「ホスト ページからダイアログ ボックスを開く [」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page))。

したがって、このメソッドにルートを渡した場合、実際には SPA は使用できないので、同じ SPA のインスタンスが 2 `displayDialogAsync` *つ必要です*。 さらに、作業ウィンドウ インスタンス内のコードの多くはこのインスタンスでは決して使用されません。また、ダイアログ ボックス インスタンス内のコードの多くが、そのインスタンスでは決して使用されません。 同じバンドルに 2 つの SPA があるようなものです。

#### <a name="microsoft-recommendations"></a>Microsoft の推奨事項

クライアント側ルートをメソッドに渡す代わりに、次のいずれかを `displayDialogAsync` 実行することをお勧めします。

* ダイアログ ボックスで実行するコードが十分に複雑な場合は、2 つの異なる SPA を明示的に作成します。つまり、同じドメインの異なるフォルダーに 2 つの SPA があります。 1 つの SPA がダイアログ ボックスで実行され、もう一方の SPA がダイアログ ボックスのホスト ページで `displayDialogAsync` 実行されます。 
* ほとんどのシナリオでは、ダイアログ ボックスで必要なのは単純なロジックのみです。 このような場合、SPA のドメインに埋め込まれたり参照された JavaScript を使用して単一の HTML ページをホストすることで、プロジェクトが大幅に簡略化されます。 ページの URL を `displayDialogAsync` メソッドに渡します。 つまり、単一ページ アプリの文字どおりの考え方から離れたものになっています。このダイアログ API を使用している場合、SPA のインスタンスは実際には 1 つOffice行う必要があります。
