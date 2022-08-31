---
title: Office ダイアログ API のベスト プラクティスとルール
description: 単一ページ アプリケーション (SPA) のベスト プラクティスなど、Office ダイアログ API のルールとベスト プラクティスを提供します。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca50e637d4b6557f508c682d2c3219f4f7dedca7
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464840"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office ダイアログ API のベスト プラクティスとルール

この記事では、ダイアログの UI を設計し、単一ページ アプリケーション (SPA) で API を使用するためのベスト プラクティスなど、Office ダイアログ API のルール、ゴチャ、ベスト プラクティスについて説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、 [Office ダイアログ API の使用](dialog-api-in-office-add-ins.md)の基本について理解していることを前提にしています。
> 
> [「Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」も参照してください。

## <a name="rules-and-gotchas"></a>ルールと注意事項

- ダイアログ ボックスは、HTTP ではなく HTTPS URL にのみ移動できます。
- [displayDialogAsync](/javascript/api/office/office.ui) メソッドに渡される URL は、アドイン自体とまったく同じドメイン内にある必要があります。 サブドメインにすることはできません。 ただし、そのページに渡されるページは、別のドメイン内のページにリダイレクトできます。
- ホスト ページでは、一度に開くことができるダイアログ ボックスは 1 つだけです。 ホスト ページには、作業ウィンドウまたは関数コマンドの [関数ファイル](/javascript/api/manifest/functionfile) を指定 [できます](../design/add-in-commands.md#types-of-add-in-commands)。
- ダイアログ ボックスで呼び出すことができる Office API は 2 つだけです。
  - [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 関数。
  - `Office.context.requirements.isSetSupported` (詳細については、「 [Office アプリケーションと API 要件の指定](specify-office-hosts-and-api-requirements.md)」を参照してください)。
- [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 関数は通常、アドイン自体とまったく同じドメイン内のページから呼び出す必要がありますが、これは必須ではありません。 詳細については、「[ホスト ランタイムへのクロスドメイン メッセージング](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime)」をご覧ください。

## <a name="best-practices"></a>ベスト プラクティス

### <a name="avoid-overusing-dialog-boxes"></a>ダイアログ ボックスを過度に使用しないようにする

UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については、 [Excel アドインの JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) サンプルを参照してください。

### <a name="design-a-dialog-box-ui"></a>ダイアログ ボックス UI を設計する

ダイアログ ボックスの設計のベスト プラクティスについては、「 [Office アドインのダイアログ ボックス](../develop/dialog-api-in-office-add-ins.md)」を参照してください。

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>Office on the webでポップアップ ブロックを処理する

Office on the webの使用中にダイアログ ボックスを表示しようとすると、ブラウザーのポップアップ ブロックによってダイアログ ボックスがブロックされる可能性があります。 この場合、次のようなプロンプトがOffice on the web開きます。

![アドインがブラウザー内のポップアップ ブロックを回避するために生成できる簡単な説明と [許可] ボタンと [無視] ボタンを含むプロンプトを示すスクリーンショット](../images/dialog-prompt-before-open.png)

ユーザーが **[許可**] を選択すると、[Office] ダイアログ ボックスが開きます。 ユーザーが **[無視**] を選択した場合、プロンプトは閉じ、Office ダイアログ ボックスは開きません。 代わりに、このメソッドは `displayDialogAsync` エラー 12009 を返します。 コードでは、このエラーをキャッチし、ダイアログを必要としない別のエクスペリエンスを提供するか、アドインでダイアログを許可するように求めるメッセージをユーザーに表示する必要があります。 (12009 の詳細については、「 [displayDialogAsync からのエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください)。

何らかの理由でこの機能を無効にする場合は、コードをオプトアウトする必要があります。この要求は、メソッドに渡される [DialogOptions](/javascript/api/office/office.dialogoptions) オブジェクトを `displayDialogAsync` 使用して行われます。 具体的には、オブジェクトに含める `promptBeforeOpen: false`必要があります。 このオプションが false に設定されている場合、Office on the webはアドインがダイアログを開くことを許可するようにユーザーに求めず、Office ダイアログは開かなくなります。

### <a name="do-not-use-the-_host_info-value"></a>ホスト\_情報の値を\_使用しない

Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 カスタム クエリ パラメーターがある場合は、その後に追加されます。 ダイアログ ボックスが移動する後続の URL には追加されません。 Microsoft は、この値の内容を変更したり、完全に削除したりして、コードで読み取らないようにすることができます。 同じ値がダイアログ ボックスのセッション ストレージ (つまり [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

### <a name="open-another-dialog-immediately-after-closing-one"></a>ダイアログを閉じた直後に別のダイアログを開く

特定のホスト ページから複数のダイアログを開く必要がないため、コードは別のダイアログを開くために呼び出す前に、開いているダイアログで [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) を呼び出 `displayDialogAsync` す必要があります。 メソッドは `close` 非同期です。 このため、呼び出しの直後に呼び出 `displayDialogAsync` すと、Office が 2 番目の `close`ダイアログを開こうとしたときに、最初のダイアログが完全に閉じなかった可能性があります。 その場合、Office は [12007](dialog-handle-errors-events.md#12007) エラーを返します。"このアドインには既にアクティブなダイアログがあるため、操作は失敗しました。

このメソッドは`close`コールバック パラメーターを受け入れません。また、Promise オブジェクトは返されないため、キーワードまたは`then`メソッドで待機`await`することはできません。 このため、ダイアログを閉じた直後に新しいダイアログを開く必要がある場合は、次の手法をお勧めします。コードをカプセル化して関数で新しいダイアログを開き、戻り値の `displayDialogAsync` 呼び出しが返された場合に関数を再帰的に呼び出すように関数を設計します `12007`。 次に例を示します。

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

または、 [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) メソッドを使用して 2 番目のダイアログを開こうとする前に、コードを強制的に一時停止することもできます。 次に例を示します。

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>SPA で Office ダイアログ API を使用するためのベスト プラクティス

アドインでクライアント側ルーティングを使用する場合は、通常、シングルページ アプリケーション (SPA) のように、別の HTML ページの URL ではなく [、displayDialogAsync](/javascript/api/office/office.ui) メソッドにルートの URL を渡すオプションがあります。 *以下に示す理由から、これを行うことをお勧めします。*

> [!NOTE]
> この記事は、Express ベースの Web アプリケーションなど、 *サーバー側* のルーティングには関係ありません。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SPA と Office ダイアログ API に関する問題

Office ダイアログ ボックスは、JavaScript エンジンの独自のインスタンスを持つ新しいウィンドウに表示されるため、完全な実行コンテキストになります。 ルートを渡すと、ベース ページとその初期化コードとブートストラップ コードがすべてこの新しいコンテキストで再度実行され、すべての変数がダイアログ ボックスの初期値に設定されます。 そのため、この手法は、アプリケーションの 2 番目のインスタンスをボックス ウィンドウでダウンロードして起動します。これは、SPA の目的を部分的に打ち負かします。 さらに、ダイアログ ボックス ウィンドウで変数を変更するコードでは、同じ変数の作業ウィンドウのバージョンは変更されません。 同様に、ダイアログ ボックス ウィンドウには独自のセッション ストレージ ( [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) があり、作業ウィンドウのコードからはアクセスできません。 ダイアログ ボックスと、呼び出されたホスト ページ `displayDialogAsync` は、サーバーに対して 2 つの異なるクライアントのように見えます。 (ホスト ページの概要については、「ホスト ページ [からダイアログ ボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください)。

そのため、メソッドにルートを `displayDialogAsync` 渡した場合、実際には SPA はありません。 *同じ SPA のインスタンスが 2 つあります*。 さらに、作業ウィンドウ インスタンス内のコードの多くは、そのインスタンスでは使用されません。ダイアログ ボックス インスタンス内のコードの多くは、そのインスタンスでは使用されません。 同じバンドルに 2 つの SPA があるようなものです。

#### <a name="microsoft-recommendations"></a>Microsoft の推奨事項

クライアント側のルートをメソッドに `displayDialogAsync` 渡す代わりに、次のいずれかを実行することをお勧めします。

* ダイアログ ボックスで実行するコードが十分に複雑な場合は、2 つの異なる SPA を明示的に作成します。つまり、同じドメインの異なるフォルダーに 2 つの SPA があります。 1 つの SPA はダイアログ ボックスで実行され、もう 1 つはダイアログ ボックスのホスト ページで実行されます。ここで `displayDialogAsync` 呼び出されました。 
* ほとんどのシナリオでは、ダイアログ ボックスには単純なロジックのみが必要です。 このような場合は、SPA のドメインに埋め込みまたは参照される JavaScript を含む単一の HTML ページをホストすることで、プロジェクトが大幅に簡素化されます。 ページの URL を `displayDialogAsync` メソッドに渡します。 これは、単一ページ アプリのリテラルアイデアから逸脱していることを意味します。Office ダイアログ API を使用している場合、実際には SPA のインスタンスが 1 つもありません。
