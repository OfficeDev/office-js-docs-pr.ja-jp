---
title: Office ダイアログ API のベスト プラクティスとルール
description: 単一ページ アプリケーション (SPA) のベスト プラクティスなどOffice API のルールとベスト プラクティスを提供します。
ms.date: 07/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: 773edd6b041ad6e49b479b3705ebcdea1875e561
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743497"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office ダイアログ API のベスト プラクティスとルール

この記事では、Office ダイアログ API のルール、ゴッチャ、ベスト プラクティスについて説明します。たとえば、ダイアログの UI を設計し、単一ページ アプリケーション (SPA) で API を使用するためのベスト プラクティスを含む。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、Office ダイアログ API の使用の基本について理解[している](dialog-api-in-office-add-ins.md)必要があります。
> 
> 「エラーと[イベントの処理とエラーの処理」Office参照してください](dialog-handle-errors-events.md)。

## <a name="rules-and-gotchas"></a>ルールと注意事項

- ダイアログ ボックスは HTTP ではなく HTTPS URL にのみ移動できます。
- [displayDialogAsync](/javascript/api/office/office.ui) メソッドに渡される URL は、アドイン自体とまったく同じドメインにある必要があります。 サブドメインにすることはできません。 ただし、そのページに渡されたページは、別のドメインのページにリダイレクトできます。
- アドイン コマンドの作業ウィンドウまたは UI レス関数ファイルを使用できるホスト ウィンドウ[](../reference/manifest/functionfile.md)では、一度に開くことができるダイアログ ボックスは 1 つのみです。
- ダイアログ ボックスOffice 2 つの API のみを呼び出します。
  - [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 関数。
  - `Office.context.requirements.isSetSupported`(詳細については、「アプリケーションと [API 要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。
- [messageParent 関数](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1))は、通常、アドイン自体とまったく同じドメイン内のページから呼び出す必要がありますが、これは必須ではありません。 詳細については、「[ホスト ランタイムへのクロスドメイン メッセージング](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime)」をご覧ください。

## <a name="best-practices"></a>ベスト プラクティス

### <a name="avoid-overusing-dialog-boxes"></a>ダイアログ ボックスの使い過ぎを回避する

UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については、「[JavaScript SalesTracker Excelを](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)参照してください。

### <a name="design-a-dialog-box-ui"></a>ダイアログ ボックスの UI を設計する

ダイアログ ボックス設計のベスト プラクティスについては、「ダイアログ ボックス」を参照[Officeアドインを参照してください](../design/dialog-boxes.md)。

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>ポップアップ ブロックを処理するには、Office on the web

ブラウザーを使用している間にダイアログ Office on the webを表示しようとすると、ブラウザーのポップアップ ブロッカーがダイアログ ボックスをブロックする可能性があります。 この場合、次Office on the webのようなプロンプトが表示されます。

![ブラウザー内ポップアップ ブロックを回避するためにアドインが生成できる簡単な説明と [許可] ボタンと [無視] ボタンを含むプロンプトを示すスクリーンショット](../images/dialog-prompt-before-open.png)

ユーザーが [許可] を **選択すると**、[Office] ダイアログ ボックスが開きます。 ユーザーが [無視] を **選択** すると、プロンプトが閉じOfficeダイアログ ボックスが開かれません。 代わりに、メソッドは `displayDialogAsync` エラー 12009 を返します。 コードは、このエラーをキャッチし、ダイアログを必要としない代替エクスペリエンスを提供するか、アドインがダイアログを許可する必要があるというメッセージをユーザーに表示する必要があります。 (12009 の詳細については、「 [displayDialogAsync からのエラー」を参照](dialog-handle-errors-events.md#errors-from-displaydialogasync)してください)。

何らかの理由でこの機能をオフにする場合は、コードをオプトアウトする必要があります。この要求は、メソッドに渡 [される DialogOptions](/javascript/api/office/office.dialogoptions) オブジェクトを使用して行 `displayDialogAsync` います。 具体的には、オブジェクトに . を含める必要があります `promptBeforeOpen: false`。 このオプションを false に設定すると、Office on the webアドインがダイアログを開くことを許可するように求めるメッセージが表示され、Officeダイアログが開かれません。

### <a name="do-not-use-the-_host_info-value"></a>hostinfo 値を \_使用\_しない

Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 カスタム クエリ パラメーターがある場合は、その後に追加されます。 ダイアログ ボックスが移動する後続の URL には追加されません。 Microsoft は、この値の内容を変更したり、完全に削除したりする場合があります。そのため、コードで読み取る必要はありません。 同じ値がダイアログ ボックスのセッション ストレージ ( [Window.sessionStorage プロパティ) に追加](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

### <a name="open-another-dialog-immediately-after-closing-one"></a>1 つを閉じるとすぐに別のダイアログを開く

特定のホスト ページから複数のダイアログを開く必要はないので、別のダイアログを開く前に、開いているダイアログで [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) `displayDialogAsync` を呼び出す必要があります。 メソッド `close` は非同期です。 このため、呼び`displayDialogAsync``close`出しの直後に呼び出した場合、2 番目のダイアログを開Officeが完全に閉じない可能性があります。 この場合、Office [12007](dialog-handle-errors-events.md#12007) エラーが返されます。"このアドインには既にアクティブなダイアログが含まれるため、操作は失敗しました。

メソッド `close` はコールバック パラメーターを受け入れないので、Promise `await` オブジェクトを返すので、キーワードまたはメソッドで待つ必要 `then` はありません。 このため、`displayDialogAsync``12007`ダイアログを閉じる直後に新しいダイアログを開く必要がある場合は、メソッドで新しいダイアログを開くコードをカプセル化し、戻り値の呼び出しが発生した場合にメソッドを再帰的に呼び出すメソッドを設計する必要がある場合は、次の方法をお勧めします。 次に例を示します。

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

または、 [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) メソッドを使用して 2 番目のダイアログを開く前に、コードを強制的に一時停止できます。 次に例を示します。

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>SPA で Office API を使用するためのベスト プラクティス

アドインがクライアント側ルーティングを使用する場合は、通常、単一ページ アプリケーション (SPA) のように、別の HTML ページの URL ではなく、ルートの URL を [displayDialogAsync](/javascript/api/office/office.ui) メソッドに渡すオプションがあります。 *以下に示す理由により、これを行うのはお勧めしません。*

> [!NOTE]
> この記事は、Express *ベース* の Web アプリケーションなど、サーバー側のルーティングには関係ありません。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SPA とダイアログ API のOffice問題

[Office] ダイアログ ボックスは、JavaScript エンジンの独自のインスタンスを持つ新しいウィンドウに表示され、それ故に、完全な実行コンテキストになります。 ルートを渡した場合、基本ページとその初期化コードとブートストラップ コードはすべて、この新しいコンテキストで再び実行され、すべての変数はダイアログ ボックスの初期値に設定されます。 したがって、この手法では、アプリケーションの 2 番目のインスタンスがボックス ウィンドウでダウンロードおよび起動され、SPA の目的の一部が打ち負かされます。 さらに、ダイアログ ボックス ウィンドウで変数を変更するコードでは、同じ変数の作業ウィンドウ バージョンは変更されません。 同様に、ダイアログ ボックス ウィンドウには独自のセッション ストレージ ( [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) があります。これは作業ウィンドウ内のコードからアクセスできません。 ダイアログ ボックスと、呼び出 `displayDialogAsync` されたホスト ページは、サーバーに対して 2 つの異なるクライアントのように見えます。 (ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)してください)。

したがって、メソッドにルート `displayDialogAsync` を渡した場合、SPA は実際には使用しないので、同じ SPA のインスタンスが 2 *つ必要になります*。 さらに、作業ウィンドウ インスタンス内のコードの多くが、そのインスタンスでは使用されません。ダイアログ ボックス インスタンス内のコードの多くは、そのインスタンスでは使用されません。 同じバンドルに 2 つの SPA があるようなものです。

#### <a name="microsoft-recommendations"></a>Microsoft の推奨事項

クライアント側ルートをメソッドに `displayDialogAsync` 渡す代わりに、次のいずれかを実行することをお勧めします。

* ダイアログ ボックスで実行するコードが十分に複雑な場合は、2 つの異なる SPA を明示的に作成します。つまり、同じドメインの異なるフォルダーに 2 つの SPA があります。 1 つの SPA はダイアログ ボックスで実行され、 `displayDialogAsync` もう 1 つはダイアログ ボックスのホスト ページで呼び出されました。 
* ほとんどのシナリオでは、ダイアログ ボックスで必要なのは単純なロジックのみです。 このような場合、SPA のドメインに JavaScript が埋め込まれているか参照されている単一の HTML ページをホストすることで、プロジェクトが大幅に簡略化されます。 ページの URL を `displayDialogAsync` メソッドに渡します。 つまり、単一ページ アプリの文字通りの考え方から離れつきます。このダイアログ API を使用している場合、SPA のインスタンスは実際には 1 つOffice必要があります。
