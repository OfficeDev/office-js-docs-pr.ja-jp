---
title: Office アドインで Office ダイアログ API を使用する
description: Office アドインでダイアログ ボックスを作成する基本について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 947b08575d100c639a440c1ca25d45199b4507ad
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68466958"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Office アドインで Office ダイアログ API を使用する

[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。 この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。

> [!NOTE]
> ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)」を参照してください。 現在、Dialog API は Excel、PowerPoint、Word でサポートされています。 Outlook のサポートは、さまざまなメールボックス要件セット&mdash;に含まれています。詳細については、API リファレンスを参照してください。

ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。 詳細については、この記事をよく読んだ *後* で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。

作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。

- 作業ウィンドウで直接開けないサインイン ページを表示します。
- アドインでの作業用に画面領域を広げる (あるいは全画面表示)。
- ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。

> [!NOTE]
> UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については、 [Excel アドインの JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) サンプルを参照してください。

次の画像は、ダイアログ ボックスの例を示します。

![Word の前に 3 つのサインイン オプションが表示されたダイアログ。](../images/auth-o-dialog-open.png)

ダイアログ ボックスが常に画面の中央に開くことに注意してください。 ユーザーはダイアログ ボックスの移動とサイズ変更ができます。 ウィンドウは *変更されません*。ユーザーは引き続き Office アプリケーションのドキュメントと作業ウィンドウのページの両方を操作できます (存在する場合)。

## <a name="open-a-dialog-box-from-a-host-page"></a>ホスト ページからダイアログ ボックスを開く

Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。

ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。 このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。 たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。

ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。 この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。 次のコードは簡単な例です。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
>
> - URL には HTTP **S** プロトコルを使用します。 これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。
> - ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](/javascript/api/manifest/functionfile)にすることができます。 ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。

> [!IMPORTANT]
> ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。 `displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。 サブドメインを含む、フル ドメインが一致している必要があります。

最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。 また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。

既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。 使用するその他の `displayDialogAsync`サンプルについては、「 [サンプル](#samples)」を参照してください。

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。 別のダイアログ ボックスを開こうとすると、エラーが発生します。 たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くできません。 ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。 これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。 詳細については、「[displayDialogAsync のエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください。

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Office on the web のパフォーマンス オプションを利用する

`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。 このプロパティを `true` に設定し、Office on the web で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。 次に例を示します。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。 アドインがOffice on the webで実行されていない場合、`displayInIframe`アドインは無視されます。

> [!NOTE]
> ダイアログ ボックスが iframe で開けないページにリダイレクトされる場合 **は使用**`displayInIframe: true`しないでください。 たとえば、Google や Microsoft アカウントなど、多くの一般的な Web サービスのサインイン ページを iframe で開くこともできます。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>ダイアログ ボックスからホスト ページに情報を送信する

> [!NOTE]
>
> - わかりやすくするために、このセクションではメッセージ ターゲットをホスト *ページ* と呼びますが、厳密に言えば、メッセージは作業ウィンドウ (または [関数ファイル](/javascript/api/manifest/functionfile)をホストしているランタイム) の [ランタイム](../testing/runtimes.md)に移動します。 この違いは、クロスドメイン メッセージングの場合にのみ重要です。 詳細については、「[ホスト ランタイムへのクロスドメイン メッセージング](#cross-domain-messaging-to-the-host-runtime)」をご覧ください。
> - Office JavaScript API ライブラリがページに読み込まれていない限り、ダイアログ ボックスは作業ウィンドウのホスト ページと通信できません。 (Office JavaScript API ライブラリを使用するページと同様に、ページのスクリプトはアドインを初期化する必要があります。 詳細については、「 [Office アドインを初期化する](initialize-add-in.md)」を参照してください。

ダイアログ ボックスのコードでは、 [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 関数を使用して、ホスト ページに文字列メッセージを送信します。 文字列には、単語、文、XML BLOB、文字列化された JSON など、文字列にシリアル化したり、文字列にキャストしたりできるものを指定できます。 次に例を示します。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
>
> - この `messageParent` 関数は、ダイアログ ボックスで呼び出すことができる 2 つの Office JS API *の 1 つです* 。
> - ダイアログ ボックスで呼び出すことができる他の JS API は `Office.context.requirements.isSetSupported`. 詳細については、「 [Office アプリケーションと API 要件の指定](specify-office-hosts-and-api-requirements.md)」を参照してください。 ただし、ダイアログ ボックスでは、この API はボリューム ライセンスの永続Outlook 2016 (つまり MSI バージョン) ではサポートされていません。

次の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

ホスト ページは、メッセージを受信するように構成する必要があります。 これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメーターを追加します。 コールバックはハンドラーを `DialogMessageReceived` イベントに割り当てます。 次に例を示します。

```js
let dialog;
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
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - `dialog` 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。

`DialogMessageReceived` イベントのハンドラーの簡単な例を次に示します。

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office は `arg` オブジェクトをハンドラーに渡します。 その `message` プロパティは、ダイアログ ボックスの `messageParent` 呼び出しによって送信される文字列です。 この例では、Microsoft アカウントや Google などのサービスからユーザーのプロファイルを文字列化して表現しているため `JSON.parse`、.
> - 実装は `showUserName` 表示されません。 作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。

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

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

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

ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。 たとえば、ダイアログ ボックスでユーザーに Microsoft アカウントや Google などの ID プロバイダーへのサインインを求めるメッセージが表示された場合、ユーザーのプロファイルはメッセージとして送信されます。 認証に失敗した場合、ダイアログ ボックスは、次の例のように、エラー情報をホスト ページに送信します。

```js
if (loginSuccess) {
    const userProfile = getProfile();
    const messageObject = {messageType: "signinSuccess", profile: userProfile};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    const errorDetails = getError();
    const messageObject = {messageType: "signinFailure", error: errorDetails};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - `loginSuccess` 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。
> - 関数と`getError`関数の`getProfile`実装は表示されません。 両方の関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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
> この `showNotification` 記事で提供されているサンプル コードでは、実装は表示されません。 アドインでこの関数を実装する方法の例は、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

### <a name="cross-domain-messaging-to-the-host-runtime"></a>ホスト ランタイムへのクロスドメイン メッセージング

ダイアログが開いたら、ダイアログまたは親ランタイムがアドインのドメインから離れる可能性があります。 これらのいずれかが発生した場合、コードで親ランタイムの `messageParent` ドメインが指定されていない限り、呼び出しは失敗します。 これを行うには、[DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) パラメーターを .`messageParent` このオブジェクトには、 `targetOrigin` メッセージの送信先となるドメインを指定するプロパティがあります。 パラメーターが使用されていない場合、Office はターゲットがダイアログが現在ホストしているのと同じドメインであると見なします。

> [!NOTE]
> クロスドメイン メッセージの送信に使用するには`messageParent`[、Dialog Origin 1.1 要件セット](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)が必要です。 このパラメーターは `DialogMessageOptions` 、要件セットをサポートしていない古いバージョンの Office では無視されるため、渡してもメソッドの動作は影響を受けません。

クロスドメイン メッセージの送信に使用 `messageParent` する例を次に示します。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> このパラメーターは `DialogMessageOptions` 、2021 年 7 月 19 日にリリースされました。 その日から約 30 日間、Office on the webでは、パラメーターなしで`DialogMessageOptions`初めて`messageParent`呼び出され、親がダイアログとは異なるドメインである場合、ユーザーはターゲット ドメインへのデータの送信を承認するように求められます。 ユーザーが承認した場合、ユーザーの回答は 24 時間キャッシュされます。 同じターゲット ドメインで呼び出された `messageParent` この期間中、ユーザーに再度メッセージが表示されることはありません。

メッセージに機密データが含まれていない場合は、"" に\*設定`targetOrigin`して任意のドメインに送信できます。 次に例を示します。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> パラメーターは `DialogMessageOptions` 、2021 年半ばに必須パラメーターとしてメソッドに追加 `messageParent` されました。 メソッドを使用してクロスドメイン メッセージを送信する以前のアドインは、新しいパラメーターを使用するように更新されるまで機能しなくなりました。 アドインが更新されるまで、 *Office for Windows でのみ*、ユーザーとシステム管理者は、レジストリ設定で信頼されたドメインを指定することで、これらのアドインを引き続き機能 **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains** できます。 これを行うには、拡張子を持つ `.reg` ファイルを作成し、Windows コンピューターに保存してからダブルクリックして実行します。 このようなファイルの内容の例を次に示します。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>情報をダイアログ ボックスに渡す

アドインは、[Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) を使用して[、ホスト ページ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)からダイアログ ボックスにメッセージを送信できます。

### <a name="use-messagechild-from-the-host-page"></a>ホスト ページから使用 `messageChild()` する

Office ダイアログ API を呼び出してダイアログ ボックスを開くと、 [Dialog](/javascript/api/office/office.dialog) オブジェクトが返されます。 オブジェクトは他のメソッドによって参照されるため、 [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) メソッドよりもスコープが大きい変数に割り当てる必要があります。 次に例を示します。

```javascript
let dialog;
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

この `Dialog` オブジェクトには、文字列化されたデータを含む任意の文字列をダイアログ ボックスに送信する [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) メソッドがあります。 これにより、ダイアログ ボックスで `DialogParentMessageReceived` イベントが発生します。 次のセクションに示すように、コードでこのイベントを処理する必要があります。

ダイアログの UI が現在アクティブなワークシートと、他のワークシートに対するワークシートの位置に関連するシナリオを考えてみましょう。 次の例では、 `sheetPropertiesChanged` Excel ワークシートのプロパティをダイアログ ボックスに送信します。 この場合、現在のワークシートは "My Sheet" という名前で、ブック内の 2 番目のシートです。 データはオブジェクトにカプセル化され、文字列化されて渡 `messageChild`すことができます。

```javascript
function sheetPropertiesChanged() {
    const messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>ダイアログ ボックスで DialogParentMessageReceived を処理する

ダイアログ ボックスの JavaScript で、[UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) メソッドを使用してイベントのハンドラー`DialogParentMessageReceived`を登録します。 これは通常、次に示すように [、Office.onReady または Office.initialize 関数](initialize-add-in.md)で行われます。 (より堅牢な例については、この記事の後半で説明します)。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

次に、ハンドラーを `onMessageFromParent` 定義します。 次のコードは、前のセクションの例を続けます。 Office はハンドラーに引数を渡し、引数オブジェクトのプロパティにホスト ページの文字列が含まれていること `message` に注意してください。 この例では、メッセージがオブジェクトに再変換され、jQuery を使用してダイアログの先頭見出しを新しいワークシート名と一致するように設定します。

```javascript
function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

ハンドラーが正しく登録されていることを確認することをお勧めします。 これを行うには、メソッドにコールバックを `addHandlerAsync` 渡します。 これは、ハンドラーの登録試行が完了したときに実行されます。 ハンドラーが正常に登録されなかった場合は、ハンドラーを使用してログに記録するか、エラーを表示します。 次に例を示します。 ここで定義されていない関数であり `reportError` 、エラーをログに記録または表示します。

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

ホスト ページから複数 `messageChild` の呼び出しを行うことができますが、イベントのダイアログ ボックス `DialogParentMessageReceived` には 1 つのハンドラーしかないため、ハンドラーは条件付きロジックを使用して異なるメッセージを区別する必要があります。 これは、条件付きメッセージングの説明に従って、ダイアログ ボックスがホスト ページにメッセージを送信するときに条件付きメッセージングを構造化する方法と正確に平行な [方法で行](#conditional-messaging)うことができます。

> [!NOTE]
> 場合によっては、`messageChild`[DialogApi 1.2 要件セット](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)の一部である API がサポートされない場合があります。 親からダイアログ ボックスへのメッセージングの別の方法については、 [ホスト ページからダイアログ ボックスにメッセージを渡す別の方法](parent-to-dialog.md)について説明します。

> [!IMPORTANT]
> [DialogApi 1.2 要件セット](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)は、アドイン マニフェストのセクションでは **\<Requirements\>** 指定できません。 メソッドと要件セットのサポートのランタイム チェックで説明されているように、メソッドを使用して `isSetSupported` 、実行時に DialogApi 1.2 [のサポートを確認する必要](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)があります。 マニフェスト要件のサポートは開発中です。

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>ダイアログ ランタイムへのドメイン間メッセージング

ダイアログが開いたら、ダイアログまたは親ランタイムがアドインのドメインから離れる可能性があります。 これらのいずれかが発生した場合、コードでダイアログ ランタイムのドメインが指定されていない限り、呼び出し `messageChild` は失敗します。 これを行うには、[DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) パラメーターを .`messageChild` このオブジェクトには、 `targetOrigin` メッセージの送信先となるドメインを指定するプロパティがあります。 パラメーターが使用されていない場合、Office はターゲットが親ランタイムが現在ホストしているドメインと同じドメインであると見なします。

> [!NOTE]
> クロスドメイン メッセージの送信に使用するには`messageChild`[、Dialog Origin 1.1 要件セット](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)が必要です。 このパラメーターは `DialogMessageOptions` 、要件セットをサポートしていない古いバージョンの Office では無視されるため、渡してもメソッドの動作は影響を受けません。

クロスドメイン メッセージの送信に使用 `messageChild` する例を次に示します。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

メッセージに機密データが含まれていない場合は、"" に\*設定`targetOrigin`して任意のドメインに *送信* できます。 次に例を示します。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

ダイアログをホストしているランタイムはマニフェストのセクションに **\<AppDomains\>** アクセスできないため、 *メッセージ* が信頼されるドメインかどうかを判断するため、ハンドラーを `DialogParentMessageReceived` 使用してこれを決定する必要があります。 ハンドラーに渡されるオブジェクトには、そのプロパティとして `origin` 親で現在ホストされているドメインが含まれています。 プロパティの使用方法の例を次に示します。

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

たとえば、コードで [Office.onReady 関数または Office.initialize 関数](initialize-add-in.md) を使用して、信頼されたドメインの配列をグローバル変数に格納できます。 その `arg.origin` 後、ハンドラー内のそのリストに対してプロパティをチェックできます。

> [!TIP]
> パラメーターは `DialogMessageOptions` 、2021 年半ばに必須パラメーターとしてメソッドに追加 `messageChild` されました。 メソッドを使用してクロスドメイン メッセージを送信する以前のアドインは、新しいパラメーターを使用するように更新されるまで機能しなくなりました。 アドインが更新されるまで、 *Office for Windows でのみ*、ユーザーとシステム管理者は、レジストリ設定で信頼されたドメインを指定することで、これらのアドインを引き続き機能 **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains** できます。 これを行うには、拡張子を持つ `.reg` ファイルを作成し、Windows コンピューターに保存してからダブルクリックして実行します。 このようなファイルの内容の例を次に示します。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>ダイアログ ボックスを閉じる

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example.

```js
function closeButtonClick() {
    const messageObject = {messageType: "dialogClosed"};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

`DialogMessageReceived` のホスト ページ ハンドラーは、この例のように `dialog.close` を呼び出します  (`dialog` オブジェクトを初期化する方法を示す前述の例を参照してください)。

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>シングルページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する

Office ダイアログ API を使用する場合は、SPA およびクライアント側のルーティングを慎重に行う必要があります。 「[SPA で Office ダイアログ API を使用する場合のベスト プラクティス](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)」を参照してください。

### <a name="error-and-event-handling"></a>エラーとイベントの処理

詳細については、「[Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」を参照し ます。

## <a name="next-steps"></a>次の手順

Office ダイアログ API に関するヒントとヘスと プラクティスの詳細については、「[Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)」を参照してください。

## <a name="samples"></a>サンプル

次のすべてのサンプルを使用 `displayDialogAsync`します。 NodeJS ベースのサーバーもあれば、ASP.NET/IIS-based サーバーを持つサーバーもありますが、アドインのサーバー側の実装方法に関係なく、メソッドを使用するロジックは同じです。

**基本：**

- [Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [トレーニング コンテンツ/ビルド アドイン (複数のサンプル)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**より複雑なサンプル:**

- [Office アドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Office アドイン NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office アドイン ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office アドイン SAAS 収益化のサンプル](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook アドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook アドイン SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Outlook アドイン トークン ビューアー](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook アドインの実行可能なメッセージ](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [OneDrive への Outlook アドイン共有](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint アドイン Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel 共有ランタイム シナリオ](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Excel アドイン ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word アドイン JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word アドイン JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word アドイン AngularJS クライアント OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Office アドイン Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office アドイン OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office アドイン UX デザイン パターン コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

** も参照してください**

- [Office アドインのランタイム](../testing/runtimes.md)