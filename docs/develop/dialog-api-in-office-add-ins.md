---
title: Office アドインでダイアログ API を使用する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 65ac55baa31f298ea5a17d789ba96772c8f5bb02
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945611"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a>Office アドインでダイアログ API を使用する

[ダイアログ API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) を使用して、Office アドインでダイアログ ボックスを開くことができます。この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。

> [!NOTE]
> ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)」を参照してください。現在、ダイアログ API は Word、Excel、PowerPoint、および Outlook でサポートされています。

> ダイアログ API の主要なシナリオは、Google や Facebook などのリソースを使用して認証を有効にすることです。 アドインで Office ユーザーのデータまたは、Office 365 や OneDrive などの Microsoft Graph を使用してアクセスできるリソースのデータが必要な場合は、可能な限り、シングル サインオン API を使用することをお勧めします。 シングル サインオンの API を使用する場合、ダイアログ API を使用する必要はありません。 詳細については、「[Office アドインのシングル サインオンを有効化する](sso-in-office-add-ins.md)」を参照してください。

作業ウィンドウ、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。

- 作業ウィンドウに直接開くことができないサインイン ページを表示する。
- アドインでの作業用に画面領域を広げる (あるいは全画面表示)。
- ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。

> [!NOTE]
> UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログを開かないようにします。作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。例については、[Excel アドイン JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) のサンプルを参照してください。

次の画像は、ダイアログ ボックスの例を示します。

![アドイン コマンド](../images/auth-o-dialog-open.png)

ダイアログ ボックスが常に画面の中央に開くことに注意してください。ユーザーはダイアログ ボックスの移動とサイズ変更ができます。ウィンドウは*モードレス*です。ホスト Office アプリケーションのドキュメントの操作と、作業ウィンドウのホスト ページ (存在する場合) の操作の両方を続行できます。

## <a name="dialog-api-scenarios"></a>ダイアログ API のシナリオ

Office JavaScript API は、[Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトと [Office.context.ui 名前空間](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js)の 2 つの関数を使用する次のシナリオをサポートしています。

### <a name="open-a-dialog-box"></a>ダイアログ ボックスを開く

ダイアログ ボックスを開くには、作業ウィンドウのコードで [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) メソッドを呼び出して、開くリソースの URL を渡します。これは、通常はページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。この記事では、'ページ' または 'Web サイト' とは、ダイアログ内のリソースを意味します。次のコードは簡単な例を示しています。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - URL には HTTP**S** プロトコルを使用します。これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。
> - ドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](https://docs.microsoft.com/javascript/office/manifest/functionfile?view=office-js)にすることができます。ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。

最初のページ (または他のリソース) が読み込まれると、ユーザーは HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。

既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)

> [!NOTE]
> ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。別のダイアログ ボックスを開こうとすると、エラーが発生します。たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合には、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くことができません。ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。詳細については、「[displayDialogAsync のエラー](#errors-from-displaydialogasync)」を参照してください。

### <a name="take-advantage-of-a-performance-option-in-office-online"></a>Office Online のパフォーマンス オプションを利用する

プロパティは `displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。このプロパティを `true` に設定し、Office Online で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。次に例を示します。`displayInIframe`

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。 アドインが Office Online で実行されていない場合、`displayInIframe` は無視されます。

> [!NOTE]
> どの時点であっても、iframe で開けないページにダイアログがリダイレクトされることになる場合は、`displayInIframe: true` を使用**しないでください**。たとえば、Google や Microsoft アカウントなどの多くの一般的な Web サービスのサインイン ページは iframe で開くことができません。

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a>ダイアログ ボックスからホスト ページに情報を送信する

ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。

- ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。
- Office JavaScript ライブラリがページに読み込まれている。(Office JavaScript ライブラリを使用するすべてのページと同様に、ページのスクリプトは `Office.initialize` プロパティにメソッドを割り当てる必要があります (空のメソッドでもかまいません)。詳細については、「[アドインの初期化](understanding-the-javascript-api-for-office.md#initializing-your-add-in)」を参照してください。)

ダイアログ ページのコードは、`messageParent` 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。文字列には、単語、文、XML BLOB、文字列に変換された JSON、または文字列にシリアル化できるすべてのものを指定できます。次に例を示します。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - 関数は、ダイアログ ボックスで呼び出すことができる、*ただ* 2 つの Office API のうちの 1 つです。もう 1 つは `Office.context.requirements.isSetSupported` です。詳細は、「[Office のホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してださい。`messageParent`
> - 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。`messageParent`

次の例では、`googleProfile` は文字列に変換されたバージョンのユーザーの Google プロファイルです。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

ホスト ページは、メッセージを受信するように構成する必要があります。これを構成するには、`displayDialogAsync` の元の呼び出しにコールバック パラメーターを追加します。コールバックはハンドラーを `DialogMessageReceived` イベントに割り当てます。次に例を示します。

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
> - Office は [AsyncResult]() オブジェクトをコールバックに渡します。Office はダイアログ ボックスを開こうとした結果を表します。ただし、ダイアログ ボックスでのイベントの結果は表しません。この違いの詳細については、「[エラーとイベントの処理](#handle-errors-and-events)」セクションを参照してください。
> - の `value` プロパティは [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトに設置されます。このオブジェクトはダイアログ ボックスの実行コンテキストではなく、ホスト ページに存在します。`asyncResult`
> - はイベントを処理する関数です。任意の名前を指定できます。`processMessage`
> - 変数は、`processMessage` でも参照されるため、コールバックよりも広い範囲で宣言されます。`dialog`

イベントのハンドラーの簡単な例を次に示します。`DialogMessageReceived`

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - Office は `arg` オブジェクトをハンドラーに渡します。その `message` プロパティは、ダイアログの `messageParent` の呼び出しで送信されるブール値または文字列です。この例では、Microsoft アカウントまたは Google などのサービスからのユーザーのプロファイルの文字列に変換された表記です。このため、`JSON.parse` を含むオブジェクトに逆シリアル化されます。
> - 実装は表示されません。作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。`showUserName`

ダイアログ ボックスのユーザー操作が完了すると、次の例に示すようにメッセージ ハンドラーはダイアログ ボックスを閉じます。

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - オブジェクトは `displayDialogAsync` の呼び出しによって返されるものと同じである必要があります。`dialog`
> - の呼び出しは、直ちにダイアログ ボックスを閉じるよう Office に指示します。`dialog.close`

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

これを実行するアドインの例については、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」のサンプルを参照してください。

#### <a name="conditional-messaging"></a>条件付きのメッセージング
ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。たとえば、ユーザーに対して Microsoft アカウントまたは Google などの ID プロバイダーにサインインするよう求めるダイアログ ボックスが表示されると、ダイアログ ボックスはユーザーのプロファイルをメッセージとして送信します。認証が失敗した場合、次の例のように、ダイアログ ボックスはホスト ページにエラー情報を送信します。

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
> - 変数は、ID プロバイダーからの HTTP 応答を読み取ることによって初期化されます。`loginSuccess`
> - 関数と `getError` 関数の実装は表示されません。両方の関数はそれぞれ、クエリ パラメーターまたは HTTP 応答の本文からデータを取得します。`getProfile`
> - サインインが成功したかどうかに応じて、さまざまな種類の匿名のオブジェクトが送信されます。両方の関数に `messageType` プロパティがありますが、一方には `profile` プロパティ、もう一方には `error` プロパティがあります。

条件付きメッセージを使用するサンプルについては、次を参照してください。
- [Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

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

### <a name="closing-the-dialog-box"></a>ダイアログ ボックスを閉じる

ダイアログ ボックスを閉じるボタンをダイアログ ボックス内に実装できます。これを実行するには、ボタンのクリック イベント ハンドラーは `messageParent` を使用して、ボタンがクリックされたことをホスト ページに通知する必要があります。次に例を示します。

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

のホスト ページ ハンドラーは、この例のように `dialog.close` を呼び出します (ダイアログ オブジェクトを初期化する方法を示す、前述の例を参照してください)。`DialogMessageReceived`


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

この手法を使用するサンプルについては、「[Office アドインの UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」リポジトリの[ダイアログ ナビゲーション設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)を参照してください。

独自の終了ダイアログ UI がない場合でも、エンド ユーザーは右上隅にある **X** を選択してダイアログ ボックスを閉じることができます。この操作により `DialogEventReceived` イベントがトリガーされます。イベントがトリガーされたときに、ホスト ウィンドウに通知する必要がある場合、ホスト ウィンドウはこのイベントのハンドラーを宣言する必要があります。詳細については、「[ダイアログ ウィンドウでのエラーとイベント](#errors-and-events-in-the-dialog-window)」セクションを参照してください。

## <a name="handle-errors-and-events"></a>エラーとイベントを処理する

コードでイベントの 2 つのカテゴリを処理する必要があります。

- ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。
- ダイアログ ウィンドウでのエラーと他のイベント。

### <a name="errors-from-displaydialogasync"></a>displayDialogAsync のエラー

一般的なプラットフォームやシステムのエラーの他に、`displayDialogAsync` の呼び出しに特有の次のエラーがあります。

|コード番号|意味|
|:-----|:-----|
|12004|に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。`displayDialogAsync`|
|12005|に渡される URL には HTTP プロトコルを使用します。HTTPS が必要です。(Office の一部のバージョンでは、12004 で返されるのと同じエラー メッセージが 12005 でも返されます。)`displayDialogAsync`|
|<span id="12007">12007</span>|ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。|

が呼び出されると、常に [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) オブジェクトがコールバック関数に渡されます。呼び出しが成功した場合 (つまり、ダイアログ ウィンドウが開いた場合)、`AsyncResult` オブジェクトの `value` プロパティは [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) オブジェクトです。この例は、「[ダイアログ ボックスからホスト ページに情報を送信する](#send-information-from-the-dialog-box-to-the-host-page)」セクションで参照できます。`displayDialogAsync` への呼び出しが失敗した場合は、ウィンドウは作成されず、`AsyncResult` オブジェクトの `status` プロパティが "失敗" に設定され、オブジェクトの `error` プロパティが設定されます。`status` をテストして、エラーが発生したときに応答するコールバックを常に設定しておく必要があります。コード番号に関係なくエラー メッセージのみを報告する例については、次のコードを参照してください。`displayDialogAsync`

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === "failed") {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a>ダイアログ ウィンドウでのエラーとイベント

ダイアログ ボックス内の、コード番号で知られている 3 つのエラーとイベントによって、ホスト ページで `DialogEventReceived` イベントがトリガーされます。

|コード番号|意味|
|:-----|:-----|
|12002|以下のいずれか:<br> - `displayDialogAsync` に渡された URL にページが存在しない。<br> - `displayDialogAsync` に渡されたページが読み込まれたが、ダイアログ ボックスが見つからないか読み込むことができないページを指していたか、またはダイアログ ボックスが無効な構文を含む URL を指している。|
|12003|ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。|
|12006|ダイアログ ボックスが閉じられました。通常は、ユーザーが **X** ボタンを選択したためです。|

コードで、呼び出し内の `DialogEventReceived` イベントのハンドラーを `displayDialogAsync` に割り当てることができます。次に簡単な例を示します。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

各エラー コードのカスタム エラー メッセージを作成する `DialogEventReceived` イベントのハンドラーの例を、次に示します。

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。


## <a name="pass-information-to-the-dialog-box"></a>情報をダイアログ ボックスに渡す

ホスト ページがダイアログ ボックスに情報を渡す必要がある場合もあります。これは主に 2 つの方法で実行することができます。

- に渡される URL にクエリ パラメーターを追加します。`displayDialogAsync`
- ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。2 つのウィンドウは共通のセッション ストレージを共有しませんが、ポート番号 (存在する場合) を含む*ドメインが同じである場合*は、共通の[ローカル ストレージ](http://www.w3schools.com/html/html5_webstorage.asp)を共有します。

### <a name="use-local-storage"></a>ローカル ストレージの使用

ローカル ストレージを使用するには、次の例に示すように、`displayDialogAsync` 呼び出しの前に、コードはホスト ページで `window.localStorage` オブジェクトの `setItem` メソッドを呼び出します。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

ダイアログ ウィンドウ内のコードは、次の例に示すように、必要に応じて項目を読み取ります。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

この方法でローカル ストレージを使用するサンプル アドインについては、次を参照してください。

- [Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a>クエリ パラメーターの使用

次の例は、クエリ パラメーターを使用してデータを渡す方法を示します。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。

ダイアログ ウィンドウ内のコードは、URL を解析し、パラメーター値を読み取ります。

> [!NOTE]
> Office は、`displayDialogAsync` に渡される URL に `_host_info` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。ダイアログ ボックスが移動する先の後続の URL には追加されません)。Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。ダイアログ ボックスのセッション ストレージには、同じ値が追加されます。この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

## <a name="use-the-dialog-apis-to-show-a-video"></a>ダイアログ API を使用してビデオを表示する

ダイアログ ボックスでビデオを表示するには

1.  コンテンツのみが iframe であるページを作成します。iframe の `src` 属性はオンライン ビデオをポイントします。ビデオの URL のプロトコルは HTTP**S** である必要があります。この記事では、このページを "video.dialogbox.html" と呼びます。マークアップの例を次に示します。

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  video.dialogbox.html ページは、ホスト ページと同じドメインにある必要があります。
3.  ホスト ページで `displayDialogAsync` の呼び出しを使用して、video.dialogbox.html を開きます。
4.  ユーザーがダイアログ ボックスを閉じたときに、アドインに通知する必要がある場合は、`DialogEventReceived` イベントのハンドラーを登録して、12006 イベントを処理します。詳しくは、「[ダイアログ ウィンドウでのエラーとイベント](#errors-and-events-in-the-dialog-window)」セクションを参照してください。

ダイアログ ボックスにビデオを表示するサンプルについては、「[Office アドインの UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」リポジトリの[ビデオ プレースマット設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)を参照してください。

![アドイン ダイアログ ボックスに表示されるビデオのスクリーン ショット](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a>認証フローでダイアログ API を使用する

ダイアログ API の主要なシナリオは、Microsoft アカウント、Office 365、Google、Facebook など、iframe でサインイン ページが開かないようにするリソースまたは ID プロバイダーを使用して認証を有効にすることです。

> [!NOTE]
> このシナリオでダイアログ API を使用する際には、`displayDialogAsync` への呼び出しで `displayInIframe: true` オプションを使用*しないでください*。このオプションの詳細については、この記事の「[Office Online のパフォーマンス オプションを利用する](#take-advantage-of-a-performance-option-in-office-online)」を参照してください。

シンプルで標準的な認証フローを、次に示します。

1. ダイアログ ボックスで開く最初のページは、アドインのドメイン (つまりホスト ウィンドウのドメイン) でホストされるローカル ページ (または他のリソース) です。このページには、"*NAME-OF-PROVIDER* にサインインが可能なページにリダイレクトしていますので、お待ちください。" という簡単な UI を含めることができます。「[情報をダイアログ ボックスに渡す](#pass-information-to-the-dialog-box)」に記載されているように、このページのコードは、ダイアログ ボックスに渡される情報を使用して、ID プロバイダーのサインイン ページの URL を構築します。
2. 次に、ダイアログ ウィンドウをサインイン ページにリダイレクトします。URL には、ユーザーがサインインしたらダイアログ ウィンドウを特定のページにリダイレクトするように ID プロバイダーに指示するクエリ パラメーターが含まれています。この記事では、このページを "redirectPage.html" と呼びます。(*このページはホスト ウィンドウと同じドメイン内のページにする必要があります*。これは、ダイアログ ウィンドウがサインイン試行の結果を渡す唯一の方法が `messageParent` の呼び出しを使用することであるためです。この呼び出しは、ホスト ウィンドウと同じドメインのページでしか行うことができません。)
2. ID プロバイダーのサービスは、ダイアログ ウィンドウからの着信 GET 要求を処理します。ユーザーが既にログオンしている場合は、直ちにウィンドウを redirectPage.html にリダイレクトして、ユーザー データをクエリ パラメーターとして含めます。ユーザーがまだサインインしていない場合は、プロバイダーのサインイン ページがウィンドウに表示され、ユーザーがサインインします。ほとんどのプロバイダーでは、ユーザーが正常にサインインできない場合、プロバイダーはダイアログ ウィンドウにエラー ページを表示して、redirectPage.html にはリダイレクトしません。ユーザーは隅にある **X** を選択して、ウィンドウを閉じる必要があります。ユーザーが正常にサインインした場合は、ダイアログ ウィンドウが redirectPage.html にリダイレクトされ、ユーザー データがクエリ パラメーターとして含まれます。
3. edirectPage.html ページが開くと、`messageParent` を呼び出して、成功または失敗をホスト ページに報告し、また必要に応じて、ユーザー データまたはエラー データも報告します。
4. イベントがホスト ページで発生し、そのハンドラーはダイアログ ウィンドウを閉じ、メッセージの他の処理を必要に応じて実行します。`DialogMessageReceived`

このパターンを使用するサンプル アドインについては、以下を参照してください。

- [PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart):ダイアログ ウィンドウで最初に開かれるリソースは、独自のビューがないコントローラーのメソッドです。これは次に、Office 365 のサインイン ページにリダイレクトします。
- [Office アドイン Office 365 のクライアント認証 AngularJS 用](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth):ダイアログ ウィンドウで最初に開かれるリソースは、ページです。

#### <a name="support-multiple-identity-providers"></a>複数の ID プロバイダーのサポート

アドインによってユーザーが Microsoft アカウント、Google、Facebook などのプロバイダーを選択できる場合は、ユーザーがプロバイダーを選択するための UI を提供するローカルの最初のページ (前述のセクションを参照) が必要です。選択すると、サインイン URL とその URL へのリダイレクトの構築がトリガーされます。

このパターンを使用するサンプルについては、「[Auth0 サービスを使用してソーシャル ログインを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-Auth0)」を参照してください。

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>外部リソースへのアドインの承認

最新の Web では、Web アプリケーションはユーザと同等の重要なセキュリティ プリンシパルであり、アプリケーションは Office 365、Google+、Facebook、LinkedIn などのオンライン リソースに対する独自の ID とアクセス許可を持っています。アプリケーションは、展開前にリソース プロバイダーに登録されます。登録には以下が含まれています。

- アプリケーションが必要とする、ユーザーのリソースへのアクセス許可の一覧。
- アプリケーションがサービスにアクセスするときに、リソース サービスがアクセス トークンを返す宛先の URL。  

リソース サービスのユーザーのデータにアクセスするアプリケーションでユーザーが関数を呼び出すと、ユーザーはサービスにサインインするように求められ、アプリケーションが必要とするユーザーのリソースへのアクセス許可をアプリケーションに付与するように求められます。次に、サービスはサインイン ウィンドウを既に登録済みの URL にリダイレクトし、アクセス トークンを渡します。アプリケーションはアクセス トークンを使用して、ユーザーのリソースにアクセスします。

ユーザーのサインイン用に示されているフローと類似したフローを使用すると、ダイアログ API を使用してこのプロセスを管理できます。違いは次の点のみです。

- ユーザーがアプリケーションが必要とするアクセス許可をアプリケーションに付与したことがない場合は、サインインすると、ユーザーに対してこれを実行するよう求めるメッセージがダイアログ ボックスに表示されます。
- ダイアログ ウィンドウは、`messageParent` を使用して文字列に変換されたアクセス トークンを送信するか、またはホスト ウィンドウがアクセス トークンを取得できる場所にアクセス トークンを格納することで、アクセス トークンをホスト ウィンドウに送信します。トークンには制限時間がありますが、制限時間内であれば追加のメッセージを表示することなく、ホスト ウィンドウはトークンを使用して、ユーザーのリソースに直接アクセスできます。

これを実現するため、次のサンプルはダイアログ API を使用します。
- [PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - データベースにアクセス トークンを格納します。
- [OAuth.io サービスを使用して大手のオンライン サービスへのアクセスを簡略化する Office アドイン](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

アドインにおける認証と承認の詳細については、以下を参照してください。
- [Office アドインで外部サービスを承認する](auth-external-add-ins.md)
- [Office の JavaScript API ヘルパーのライブラリ](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>単一ページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する

単一ページ アプリケーションが通常使用するように、アドインがクライアント側ルーティングを使用している場合は、HTML の完了ページと個別ページの URL の代わりに、ルートの URL を [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) メソッドに渡すこともできます。

> [!IMPORTANT]
>ダイアログ ボックスは、独自の実行コンテキストを含む新しいウィンドウ内にあります。ルートを渡すと、ダイアログ ウィンドウで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。この手法により、ダイアログ ウィンドウで、アプリケーションの 2 番目のインスタンスが起動します。ダイアログ ウィンドウ内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。同様に、ダイアログ ウィンドウには、それ自体にセッション ストレージがあり、作業ウィンドウからコードでそこにアクセスすることはできません。
