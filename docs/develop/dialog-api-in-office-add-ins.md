---
title: Office アドインで Office ダイアログ API を使用する
description: アドインでダイアログ ボックスを作成する基本Office説明します。
ms.date: 08/27/2021
localization_priority: Normal
ms.openlocfilehash: 6e87ddfc6c29e74a578d399116df5df9b364028f
ms.sourcegitcommit: 3287eb4588d0af47f1ab8a59882bcc3f585169d8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2021
ms.locfileid: "58863535"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Office アドインで Office ダイアログ API を使用する

[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。 この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。

> [!NOTE]
> ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)」を参照してください。 ダイアログ API は、現在、ユーザー、Excel、PowerPoint、および Word でサポートされています。 Outlookサポートは、さまざまなメールボックス要件セットに含まれています。詳細については &mdash; 、API リファレンスを参照してください。

ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。 詳細については、この記事をよく読んだ *後* で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。

作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。

- 作業ウィンドウで直接開くことができませんサインイン ページを表示します。
- アドインでの作業用に画面領域を広げる (あるいは全画面表示)。
- ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。

> [!NOTE]
> UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については、「Excel [JavaScript SalesTracker サンプル」を参照](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)してください。

次の画像は、ダイアログ ボックスの例を示します。

![Word の前に 3 つのサインイン オプションが表示されたダイアログを示すスクリーンショット。](../images/auth-o-dialog-open.png)

ダイアログ ボックスが常に画面の中央に開くことに注意してください。 ユーザーはダイアログ ボックスの移動とサイズ変更ができます。 ウィンドウは *非* モーダルです。ユーザーは、Office アプリケーション内のドキュメントと作業ウィンドウ内のページの両方 (ある場合) の操作を続行できます。

## <a name="open-a-dialog-box-from-a-host-page"></a>ホスト ページからダイアログ ボックスを開く

Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。

ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。 このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。 たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。

ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。 この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。 次のコードは簡単な例です。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - この URL には HTTP **S** プロトコルを使用します。これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページで必須です。
> - ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](../reference/manifest/functionfile.md)にすることができます。 ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。

> [!IMPORTANT]
> ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。 `displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。 サブドメインを含む、フル ドメインが一致している必要があります。

最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。 また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。

既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。 使用するその他のサンプルについては `displayDialogAsync` [、「Samples」を参照してください](#samples)。

全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)

> [!NOTE]
> ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。 別のダイアログ ボックスを開こうとすると、エラーが発生します。 たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合、作業ウィンドウ内の別のページから 2 番目のダイアログ ボックスを開くことができません。 ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。 これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。 詳細については、「[displayDialogAsync のエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください。

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Office on the web のパフォーマンス オプションを利用する

`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。 このプロパティを `true` に設定し、Office on the web で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。 次に例を示します。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。 アドインがアプリ内で実行されていない場合Office on the web無視 `displayInIframe` されます。

> [!NOTE]
> ダイアログ ボックス **が** iframe で開くことができませんページにリダイレクトされる場合は、使用 `displayInIframe: true` する必要があります。 たとえば、Google や Microsoft アカウントなど、多くの一般的な Web サービスのサインイン ページを iframe で開くことができません。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>ダイアログ ボックスからホスト ページに情報を送信する

> [!NOTE]
>
> - わかりやすくするために、このセクションでは、メッセージ ターゲットをホスト ページと呼び出しますが、厳密に言えば、メッセージは作業ウィンドウ (または関数ファイルをホストしている [ランタイム)](../reference/manifest/functionfile.md)の *JavaScript* ランタイムに移動します。 この違いは、クロスドメイン メッセージングの場合にのみ重要です。 詳細については、「[ホスト ランタイムへのクロスドメイン メッセージング](#cross-domain-messaging-to-the-host-runtime)」をご覧ください。
> - JavaScript API ライブラリがページに読み込まれている場合をOffice、作業ウィンドウのホスト ページと通信できません。 (JavaScript API ライブラリのOfficeページと同様に、ページのスクリプトでアドインを初期化する必要があります。 詳細については、「アドイン[の初期化」Officeを参照してください](initialize-add-in.md)。

ダイアログ ボックスのコードでは [、messageParent 関数を使用](/javascript/api/office/office.ui#messageParent_message__messageOptions_) してホスト ページに文字列メッセージを送信します。 文字列には、単語、文、XML BLOB、文字列化された JSON など、文字列にシリアル化したり、文字列にキャストしたりできる文字列を指定できます。 次に例を示します。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - この `messageParent` 関数は、ダイアログボックスでOfficeできる 2 つの JS API の 1 つのみです。
> - ダイアログ ボックスで呼び出す他の JS API はです `Office.context.requirements.isSetSupported` 。 詳細については、「アプリケーションと[API 要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。 ただし、ダイアログ ボックスでは、この API は 1 回Outlook 2016 (MSI バージョン) ではサポートされていません。

次の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

ホスト ページは、メッセージを受信するように構成する必要があります。 これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメーターを追加します。 コールバックはハンドラーを `DialogMessageReceived` イベントに割り当てます。 次に例を示します。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
>
> - Office は [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトをコールバックに渡します。 Office はダイアログ ボックスを開こうとした結果を表します。 ただし、ダイアログ ボックスでのイベントの結果は表しません。 この違いの詳細については、「[エラーとイベントの処理](dialog-handle-errors-events.md)」を参照してください。
> - `asyncResult` の `value` プロパティは [Dialog](/javascript/api/office/office.dialog) オブジェクトに設置されます。このオブジェクトはダイアログ ボックスの実行コンテキストではなく、ホスト ページに存在します。
> - `processMessage` はイベントを処理する関数です。任意の名前を指定できます。
> - `dialog` 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。

`DialogMessageReceived` イベントのハンドラーの簡単な例を次に示します。

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office は `arg` オブジェクトをハンドラーに渡します。 プロパティ `message` は、ダイアログ ボックスの呼び出しによって送信 `messageParent` される文字列です。 この例では、Microsoft アカウントや Google などのサービスからユーザーのプロファイルを文字列で表すので、オブジェクトに逆シリアル化されます `JSON.parse` 。
> - 実装 `showUserName` は表示されません。 作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。

ダイアログ ボックスのユーザー操作が完了すると、次の例に示すようにメッセージ ハンドラーはダイアログ ボックスを閉じます。

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
>
> - `dialog` オブジェクトは `displayDialogAsync` の呼び出しによって返されるものと同じである必要があります。
> - `dialog.close` の呼び出しは、直ちにダイアログ ボックスを閉じるよう Office に指示します。

これらの手法を使用するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

メッセージを受信した後、アドインで作業ウィンドウの別のページを開く必要がある場合は、ハンドラーの最後の行として `window.location.replace` メソッド (または `window.location.href`) を使用できます。次に例を示します。

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

これを実行するアドインの例については、「[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」 (PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入) のサンプルを参照してください。

### <a name="conditional-messaging"></a>条件付きのメッセージング

ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。 たとえば、ダイアログ ボックスでユーザーに Microsoft アカウントや Google などの ID プロバイダーへのサインインを求めるメッセージが表示された場合、ユーザーのプロファイルがメッセージとして送信されます。 認証に失敗した場合、次の例のように、ダイアログ ボックスはエラー情報をホスト ページに送信します。

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - `loginSuccess` 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。
> - `getProfile` 関数と `getError` 関数の実装は表示されません。両方の関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。
> - サインインが成功したかどうかに応じて、さまざまな種類の匿名のオブジェクトが送信されます。両方の関数に `messageType` プロパティがありますが、一方には `profile` プロパティ、もう一方には `error` プロパティがあります。

次の例に示すように、ホスト ページのハンドラー コードは分岐に `messageType` プロパティの値を使用します。`showUserName` 関数は上記の例と同じであり、`showNotification` 関数はホスト ページの UI にエラーを表示することに注意してください。

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> 実装 `showNotification` は、この記事で提供されるサンプル コードには表示されません。 アドインでこの関数を実装する方法の例は、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

### <a name="cross-domain-messaging-to-the-host-runtime"></a>ホスト ランタイムへのクロスドメイン メッセージング

ダイアログまたは親 JavaScript ランタイム (作業ウィンドウまたは関数ファイルをホストする UI レス ランタイムのいずれか) は、ダイアログを開いた後にアドインのドメインから移動できます。 これらのいずれかのことが発生した場合、コードで親ランタイムのドメインを指定しない限り、呼び出し `messageParent` は失敗します。 これを行うには [、DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) パラメーターを呼び出しに追加します `messageParent` 。 このオブジェクトには、 `targetOrigin` メッセージを送信するドメインを指定するプロパティがあります。 パラメーターを使用しない場合、Officeは、ダイアログが現在ホストしているドメインと同じドメインである必要があります。

> [!NOTE]
> クロス `messageParent` ドメイン メッセージの送信に使用するには [、Dialog Origin 1.1 要件セットが必要です](../reference/requirement-sets/dialog-origin-requirement-sets.md)。 このパラメーターは、要件セットをサポートしていない古いバージョンの Officeでは無視されます。そのため、メソッドを渡した場合、メソッドの動作は `DialogMessageOptions` 影響を受けません。

次に、クロスドメイン メッセージを `messageParent` 送信する使用例を示します。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> この `DialogMessageOptions` パラメーターは、2021 年 7 月 19 日にリリースされました。 この日付から約 30 日後の Office on the web では、パラメーターなしで初めて呼び出され、親がダイアログとは別のドメインである場合、ユーザーはターゲット ドメインへのデータの送信を承認するように求めるメッセージが表示されます。 `messageParent` `DialogMessageOptions` ユーザーが承認した場合、ユーザーの回答は 24 時間キャッシュされます。 同じターゲット ドメインで呼び出されたこの期間中、ユーザーは再び `messageParent` 要求されません。

メッセージに機密データが含まれる場合は、任意のドメインに送信できる `targetOrigin` \* " に設定できます。 次に例を示します。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> この `DialogMessageOptions` パラメーターは、2021 年半ばに必須パラメーターとしてメソッド `messageParent` に追加されました。 メソッドでクロスドメイン メッセージを送信する古いアドインは、新しいパラメーターを使用するために更新されるまで機能しなくなりました。 アドインが更新されるまで *、Office* の Windows でのみ、ユーザーとシステム管理者は、レジストリ設定で信頼できるドメインを指定することで、これらのアドインの作業を続行できます。HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains。 **** これを行うには、拡張子を持つファイルを作成し、そのファイルを Windowsコンピューターに保存し、ダブルクリックして `.reg` 実行します。 次に、このようなファイルの内容の例を示します。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>情報をダイアログ ボックスに渡す

アドインは、Dialog.messageChild[](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)を使用して、ホスト ページからダイアログ ボックスに[メッセージを送信できます](/javascript/api/office/office.dialog#messageChild_message__messageOptions_)。

### <a name="use-messagechild-from-the-host-page"></a>ホスト `messageChild()` ページから使用する

ダイアログ ボックスを開Officeダイアログ API を呼び出す場合[、Dialog](/javascript/api/office/office.dialog)オブジェクトが返されます。 オブジェクトは他のメソッドによって参照されるので [、displayDialogAsync](/javascript/api/office/office.ui#displayDialogAsync_startAddress__callback_) メソッドよりもスコープの大きい変数に割り当てる必要があります。 次に例を示します。

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

この `Dialog` オブジェクトには [messageChild メソッドが](/javascript/api/office/office.dialog#messageChild_message__messageOptions_) 含まれており、文字列化されたデータを含む任意の文字列をダイアログ ボックスに送信します。 これにより、ダイアログ ボックス `DialogParentMessageReceived` でイベントが発生します。 次のセクションに示すように、コードでこのイベントを処理する必要があります。

ダイアログの UI が現在アクティブなワークシートに関連付け、他のワークシートを基準にしたワークシートの位置を示すシナリオを考えます。 次の例では、 `sheetPropertiesChanged` ワークシートExcelダイアログ ボックスに送信します。 この場合、現在のワークシートの名前は "My Sheet" で、ブックの 2 番目のシートです。 データはオブジェクトにカプセル化され、文字列化され、渡されます `messageChild` 。

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>ダイアログ ボックスで DialogParentMessageReceived を処理する

ダイアログ ボックスの JavaScript で、イベントのハンドラーを `DialogParentMessageReceived` [UI.addHandlerAsync メソッドに登録](/javascript/api/office/office.ui#addHandlerAsync_eventType__handler__options__callback_) します。 これは通常[、Office.onReady](initialize-add-in.md)メソッドまたは Office.initialize メソッドで行われます。次に示すようにします。 (より堅牢な例を以下に示します。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

次に、ハンドラーを定義 `onMessageFromParent` します。 次のコードは、前のセクションからこの例を続行します。 引数がOfficeハンドラーに渡され、引数オブジェクトのプロパティにホスト ページの文字列 `message` が含まれる点に注意してください。 この例では、メッセージがオブジェクトに再変換され、jQuery を使用して、新しいワークシート名と一致するダイアログの上部見出しを設定します。

```javascript
function onMessageFromParent(arg) {
    var messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

ハンドラーが正しく登録されていることを確認するベスト プラクティスです。 これを行うには、メソッドにコールバックを渡 `addHandlerAsync` します。 これは、ハンドラーの登録が完了すると実行されます。 ハンドラーが正常に登録されていない場合は、ハンドラーを使用してエラーを記録または表示します。 次に例を示します。 ここで定義 `reportError` されていない関数で、エラーをログに記録または表示する点に注意してください。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>親ページからダイアログ ボックスへの条件付きメッセージング

ホスト ページから複数の呼び出しを実行できますが、イベントのダイアログ ボックスにはハンドラーが 1 つしか存在しないので、ハンドラーは、さまざまなメッセージを区別するために条件付きロジックを使用 `messageChild` `DialogParentMessageReceived` する必要があります。 これは、「条件付きメッセージング」の説明に従って、ダイアログ ボックスがホスト ページにメッセージを送信するときに条件付きメッセージングを構成する方法と正確に並行して実行 [できます](#conditional-messaging)。

> [!NOTE]
> 場合によっては `messageChild` [、DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md)要件セットの一部である API がサポートされない場合があります。 親からダイアログ ボックスへのメッセージングの代替方法については、「ホスト ページからダイアログ ボックスにメッセージを渡す別の方法 [」を参照してください](parent-to-dialog.md)。

> [!IMPORTANT]
> DialogApi [1.2 要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md) は、アドイン マニフェストのセクション `<Requirements>` では指定できない。 [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)メソッドを使用して、実行時に DialogApi 1.2 のサポートを確認する必要があります。 マニフェスト要件のサポートは開発中です。

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>ダイアログ ランタイムへのクロスドメイン メッセージング

ダイアログまたは親 JavaScript ランタイム (作業ウィンドウまたは関数ファイルをホストする UI レス ランタイムのいずれか) は、ダイアログを開いた後にアドインのドメインから移動できます。 これらのいずれかのことが発生した場合、コードでダイアログ ランタイムのドメインを指定しない限り、呼び出し `messageChild` は失敗します。 これを行うには [、DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) パラメーターを呼び出しに追加します `messageChild` 。 このオブジェクトには、 `targetOrigin` メッセージを送信するドメインを指定するプロパティがあります。 パラメーターを使用しない場合、Officeは、親ランタイムが現在ホストしているドメインと同じドメインである必要があります。 

> [!NOTE]
> クロス `messageChild` ドメイン メッセージの送信に使用するには [、Dialog Origin 1.1 要件セットが必要です](../reference/requirement-sets/dialog-origin-requirement-sets.md)。 このパラメーターは、要件セットをサポートしていない古いバージョンの Officeでは無視されます。そのため、メソッドを渡した場合、メソッドの動作は `DialogMessageOptions` 影響を受けません。

次に、クロスドメイン メッセージを `messageChild` 送信する使用例を示します。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

メッセージに機密データが含まれる場合は、任意のドメインに送信できる `targetOrigin` \* " に設定できます。 次に例を示します。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

ダイアログをホストしている JavaScript ランタイムはマニフェストのセクションにアクセスできないので、メッセージが送信されるドメインが信頼されているかどうかを判断するため、ハンドラーを使用してこれを判断する必要があります。 `<AppDomains>`  `DialogParentMessageReceived` ハンドラーに渡されるオブジェクトには、親で現在ホストされているドメインがプロパティとして含 `origin` まれる。 プロパティの使い方の例を次に示します。

```javascript
function onMessageFromParent(arg) {
    if (arg.origin === "https://addin.fabrikam.com") {
        // process message
    } else {
        dialog.close();
        showNotification("Messages from " + arg.origin + " are not accepted.");
    }
}
```

たとえば、コードで[Office.onReady](initialize-add-in.md)メソッドまたは Office.initialize メソッドを使用して、信頼できるドメインの配列をグローバル変数に格納できます。 その `arg.origin` 後、ハンドラー内のそのリストに対してプロパティをチェックできます。

> [!TIP]
> この `DialogMessageOptions` パラメーターは、2021 年半ばに必須パラメーターとしてメソッド `messageChild` に追加されました。 メソッドでクロスドメイン メッセージを送信する古いアドインは、新しいパラメーターを使用するために更新されるまで機能しなくなりました。 アドインが更新されるまで *、Office* の Windows でのみ、ユーザーとシステム管理者は、レジストリ設定で信頼できるドメインを指定することで、これらのアドインの作業を続行できます。HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains。 **** これを行うには、拡張子を持つファイルを作成し、そのファイルを Windowsコンピューターに保存し、ダブルクリックして `.reg` 実行します。 次に、このようなファイルの内容の例を示します。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>ダイアログ ボックスを閉じる

ダイアログ ボックスを閉じるボタンをダイアログ ボックス内に実装できます。これを実行するには、ボタンのクリック イベント ハンドラーは `messageParent` を使用して、ボタンがクリックされたことをホスト ページに通知する必要があります。次に例を示します。

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

`DialogMessageReceived` のホスト ページ ハンドラーは、この例のように `dialog.close` を呼び出します  (`dialog` オブジェクトを初期化する方法を示す前述の例を参照してください)。

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

独自の終了ダイアログ UI がない場合でも、エンド ユーザーは右上隅にある **X** を選択してダイアログ ボックスを閉じることができます。 この操作により `DialogEventReceived` イベントがトリガーされます。 イベントがトリガーされたときに、ホスト ウィンドウに通知する必要がある場合、ホスト ウィンドウはこのイベントのハンドラーを宣言する必要があります。 詳細については、「[ダイアログ ボックスでのエラーとイベント](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box)」セクションを参照してください。

## <a name="advanced-topics-and-special-scenarios"></a>高度なトピックと特殊なシナリオ

### <a name="use-the-dialog-api-to-show-a-video"></a>ダイアログ API を使用してビデオを表示する

「[Office ダイアログ ボックスを使用してビデオを表示する](dialog-video.md)」を参照してください 。

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a>認証フローでダイアログ API を使用する

「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>単一ページ アプリケーションOfficeクライアント側ルーティングと一緒にダイアログ API を使用する

Office ダイアログ API を使用する場合は、SPA およびクライアント側のルーティングを慎重に行う必要があります。 「[SPA で Office ダイアログ API を使用する場合のベスト プラクティス](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)」を参照してください。

### <a name="error-and-event-handling"></a>エラーとイベントの処理

詳細については、「[Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」を参照し ます。

## <a name="next-steps"></a>次の手順

Office ダイアログ API に関するヒントとヘスと プラクティスの詳細については、「[Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)」を参照してください。

## <a name="samples"></a>サンプル

次のサンプルはすべてを使用します `displayDialogAsync` 。 ノードJS ベースのサーバーを持つサーバーや、ASP.NET/IIS ベースのサーバーを持つサーバーがありますが、アドインのサーバー側の実装方法に関係なく、メソッドを使用するロジックは同じです。

**基本:**

- [Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [トレーニング コンテンツ / アドインの構築 (複数のサンプル)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**より複雑なサンプル:**

- [Officeアドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Office アドイン NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [Officeアドイン ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Officeアドイン SAAS 収益化のサンプル](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlookアドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlookアドイン SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)
- [Outlookアドイン トークン ビューアー](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlookアドインのアクション可能なメッセージ](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Outlookアドインの共有OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPointアドイン Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel共有ランタイム シナリオ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/900b5769bca9bbcff79d6cd6106d9fcc55c70d5a/Samples/excel-shared-runtime-scenario)
- [Excelアドイン ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word アドイン JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word アドイン JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word アドイン AngularJS クライアント OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Office アドイン Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Officeアドイン の OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Officeアドイン UX デザイン パターン コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
