---
title: Office アドインで Office ダイアログ API を使用する
description: 新しいアドインでダイアログ ボックスを作成するOfficeについて説明します。
ms.date: 01/28/2021
localization_priority: Normal
ms.openlocfilehash: 9061b4c048a133572e615152d61df611e5f15068
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237866"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Office アドインで Office ダイアログ API を使用する

[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。 この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。

> [!NOTE]
> ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)」を参照してください。 ダイアログ API は現在、Excel、PowerPoint、および Word でサポートされています。 Outlook のサポートは、さまざまなメールボックス要件セット全体に含まれています。詳細については &mdash; 、API リファレンスを参照してください。

ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。 詳細については、この記事をよく読んだ *後* で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。

作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。

- 作業ウィンドウに直接開くことができないサインイン ページを表示する。
- アドインでの作業用に画面領域を広げる (あるいは全画面表示)。
- ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。

> [!NOTE]
> UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 タブ付き作業ウィンドウの例については [、Excel アドインの JavaScript SalesTracker サンプルを参照](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) してください。

次の画像は、ダイアログ ボックスの例を示します。

![Word の前に 3 つのサインイン オプションが表示されたダイアログを示すスクリーンショット](../images/auth-o-dialog-open.png)

ダイアログ ボックスが常に画面の中央に開くことに注意してください。 ユーザーはダイアログ ボックスの移動とサイズ変更ができます。 ウィンドウはモーダルではありません。ユーザーは、Office アプリケーション内のドキュメントと作業ウィンドウ内のページ (ある場合) の両方を引き続き操作できます。

## <a name="open-a-dialog-box-from-a-host-page"></a>ホスト ページからダイアログ ボックスを開く

Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。

ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。 このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。 たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。

ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。 この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。 次のコードは簡単な例を示しています。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - URL には HTTP **S** プロトコルを使用します。 これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。
> - ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](../reference/manifest/functionfile.md)にすることができます。 ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。

> [!IMPORTANT]
> ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。 `displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。 サブドメインを含む、フル ドメインが一致している必要があります。

最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。 また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。

既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。 使用するその他のサンプルについては、「 `displayDialogAsync` サンプル」を [参照してください](#samples)。

全画面表示で効率的に操作するには、両方の値を 100% に設定します。(最大有効値は 99.5% であり、最大有効値にしても、ウィンドウは移動とサイズ変更が可能です。)

> [!NOTE]
> ホスト ウィンドウから開くことができるのは、1 つのダイアログ ボックスのみです。別のダイアログ ボックスを開こうとすると、エラーが発生します。たとえば、ユーザーが作業ウィンドウからダイアログ ボックスを開いた場合には、作業ウィンドウの別のページから 2 番目のダイアログ ボックスを開くことができません。ただし、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開く場合は、選択するたびにコマンドによって新しい (ただし非表示の) HTML ファイルが開かれます。これにより、新しい (非表示) ホスト ウィンドウが作成されるため、これらの各ウィンドウは独自のダイアログ ボックスを起動できます。詳細については、「[displayDialogAsync のエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください。

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Office on the web のパフォーマンス オプションを利用する

`displayInIframe` プロパティは、`displayDialogAsync` に渡すことのできる構成オブジェクトの追加のプロパティです。 このプロパティを `true` に設定し、Office on the web で開いたドキュメントでアドインを実行している場合、ダイアログ ボックスは浮動の iframe で開き、独立したウィンドウでは開きません (この方が速く開きます)。 例を次に示します。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

既定値は `false` です。これはプロパティを完全に省略した場合と同じ状態です。 アドインが Office on the web で実行されていない場合、`displayInIframe` は無視されます。

> [!NOTE]
> どの時点であれ、iframe で開けないページにダイアログ ボックスがリダイレクトされることになる場合は、`displayInIframe: true` を使用すべきでは **ありません**。 たとえば、Google や Microsoft アカウントなど、多くの一般的な Web サービスのサインイン ページは iframe で開くことができません。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>ダイアログ ボックスからホスト ページに情報を送信する

ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。

- ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。
- ページOffice JavaScript API ライブラリが読み込まれます。 (Office JavaScript API ライブラリを使用するページと同様に、ページのスクリプトは、空のメソッドにすることもできますが、プロパティにメソッドを割り `Office.initialize` 当てる必要があります。 詳細については、「 [アドインの初期化Office」を参照してください](initialize-add-in.md))。

ダイアログ ボックスのコードは、[messageParent](/javascript/api/office/office.ui#messageparent-message-) 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。 文字列には、単語、文、XML BLOB、文字列に変換された JSON、文字列にシリアル化できるすべてのものを指定できます。 例を次に示します。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - `messageParent` 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。
> - この `messageParent` 関数は、ダイアログボックスOffice呼び出し可能な 2 つの JS API の 1 つのみです。
> - ダイアログ ボックスで呼び出し可能な他の JS API は次の機能です `Office.context.requirements.isSetSupported` 。 詳細については、「アプリケーションと [API の要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。 ただし、ダイアログ ボックスでは、この API は Outlook 2016 の 1 回購入 (つまり MSI バージョン) ではサポートされていません。

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
> - Office は `arg` オブジェクトをハンドラーに渡します。 その `message` プロパティは、ダイアログ ボックスの `messageParent` の呼び出しで送信されるブール値または文字列です。 この例では、Microsoft アカウントや Google などのサービスからのユーザーのプロファイルを文字列で表すので、オブジェクトに逆シリアル化されます `JSON.parse` 。
> - `showUserName` 実装は表示されません。作業ウィンドウ上に個人用のウェルカム メッセージが表示される場合があります。

ダイアログ ボックスのユーザー操作が完了すると、次の例に示すようにメッセージ ハンドラーはダイアログ ボックスを閉じます。

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
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

ダイアログ ボックスから複数の `messageParent` 呼び出しを送信できますが、`DialogMessageReceived` イベントのホスト ページにあるハンドラーは 1 つのみのため、ハンドラーは条件ロジックを使用してさまざまなメッセージを区別する必要があります。 たとえば、Microsoft アカウントや Google などの ID プロバイダーにサインインするように求めるダイアログ ボックスが表示された場合、ユーザーのプロファイルがメッセージとして送信されます。 認証が失敗した場合、次の例のように、ダイアログ ボックスはホスト ページにエラー情報を送信します。

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
> `showNotification`の実装は、この記事のサンプル コードでは表示されません。 アドインでこの関数を実装する方法の例は、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

## <a name="pass-information-to-the-dialog-box"></a>情報をダイアログ ボックスに渡す

アドインは[、Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-)[](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)を使用してホスト ページからダイアログ ボックスにメッセージを送信できます。

### <a name="use-messagechild-from-the-host-page"></a>ホスト `messageChild()` ページから使用する

ダイアログ ボックスを開Officeダイアログ API を呼び出す場合 [、Dialog](/javascript/api/office/office.dialog) オブジェクトが返されます。 オブジェクトは他のメソッドによって参照されるので [、displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) メソッドよりもスコープが大きい変数に割り当てる必要があります。 例を次に示します。

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

このオブジェクトには、文字列に変換されたデータを含む任意の文字列をダイアログ ボックス `Dialog` に送信する [messageChild](/javascript/api/office/office.dialog#messagechild-message-) メソッドがあります。 これにより、ダイアログ `DialogParentMessageReceived` ボックスでイベントが発生します。 次のセクションに示すように、コードでこのイベントを処理する必要があります。

ダイアログの UI が現在アクティブなワークシートに関連し、そのワークシートの位置が他のワークシートを基準にしているシナリオについて考えます。 次の例では `sheetPropertiesChanged` 、Excel ワークシートのプロパティをダイアログ ボックスに送信します。 この場合、現在のワークシートの名前は "My Sheet" で、ブックの 2 番目のシートです。 データはオブジェクトにカプセル化され、文字列化され、渡しが可能になります `messageChild` 。

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

ダイアログ ボックスの JavaScript で、イベントのハンドラーを `DialogParentMessageReceived` [UI.addHandlerAsync メソッドに登録](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) します。 これは通常、次に示すように [、Office.onReady](initialize-add-in.md)メソッドまたは Office.initialize メソッドで行われます。 (より堅牢な例を以下に示します。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

次に、ハンドラーを定義 `onMessageFromParent` します。 次のコードは、前のセクションの例を続けて示しています。 ただし、Officeハンドラーに引数を渡し、引数オブジェクトのプロパティにホスト ページの文字列 `message` が含まれる点に注意してください。 この例では、メッセージがオブジェクトに再変換され、jQuery を使用して、新しいワークシート名と一致するダイアログの上部見出しを設定します。

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

ハンドラーが正しく登録されていることを確認するベスト プラクティスです。 これを行うには、メソッドにコールバックを渡 `addHandlerAsync` します。 これは、ハンドラーの登録が完了すると実行されます。 ハンドラーが正常に登録されていない場合は、ハンドラーを使用してログを記録するか、エラーを表示します。 次に例を示します。 ここで定義 `reportError` されていない関数で、エラーをログに記録または表示します。

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

ホスト ページから複数の呼び出しを実行できますが、イベントのダイアログ ボックスにはハンドラーが 1 つしか存在しないので、ハンドラーは条件ロジックを使用して異なるメッセージを区別 `messageChild` `DialogParentMessageReceived` する必要があります。 これは、ダイアログ ボックスが条件付きメッセージングの説明に従ってホスト ページにメッセージを送信するときに、条件付きメッセージングを構成する方法と正確に同じ方法で行 [います](#conditional-messaging)。

> [!NOTE]
> 場合によっては `messageChild` [、DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md)要件セットの一部である API がサポートされていない場合があります。 親からダイアログ ボックスへのメッセージングの別の方法については、ホスト ページからダイアログ ボックスにメッセージを渡す別の方法 [で説明されています](parent-to-dialog.md)。

> [!IMPORTANT]
> DialogApi [1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) 要件セットは、アドイン マニフェストの `<Requirements>` セクションでは指定できません。 [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)メソッドを使用して、実行時に DialogApi 1.2 のサポートを確認する必要があります。 マニフェスト要件のサポートは開発中です。

## <a name="closing-the-dialog-box"></a>ダイアログ ボックスを閉じる

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

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>単一ページ アプリケーションとクライアント側ルーティングで Office ダイアログ API を使用する

Office ダイアログ API を使用する場合は、SPA およびクライアント側のルーティングを慎重に行う必要があります。 「[SPA で Office ダイアログ API を使用する場合のベスト プラクティス](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)」を参照してください。

### <a name="error-and-event-handling"></a>エラーとイベントの処理

詳細については、「[Office ダイアログ ボックスでのエラーとイベントの処理](dialog-handle-errors-events.md)」を参照し ます。

## <a name="next-steps"></a>次の手順

Office ダイアログ API に関するヒントとヘスと プラクティスの詳細については、「[Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)」を参照してください。

## <a name="samples"></a>サンプル

次のサンプルはすべて使用します `displayDialogAsync` 。 NodeJS ベースのサーバーを持つサーバーと ASP.NET/IIS-based サーバーを持つサーバーがありますが、このメソッドを使用するロジックは、アドインのサーバー側の実装方法に関係なく同じです。

**基本:**

- [Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [トレーニング コンテンツ/ アドインの作成 (いくつかのサンプル)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**より複雑なサンプル:**

- [Office アドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Office アドイン NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office SAAS 収益化のサンプル](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook アドイン Microsoft Graph ASPNET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook アドイン SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)
- [Outlook アドイン トークン ビューアー](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook アドインの操作可能なメッセージ](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [OneDrive への Outlook アドインの共有](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint アドイン Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel 共有ランタイム シナリオ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/900b5769bca9bbcff79d6cd6106d9fcc55c70d5a/Samples/excel-shared-runtime-scenario)
- [Excel アドイン ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word アドイン JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word アドイン JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word アドイン AngularJS クライアント OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Office アドイン Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office アドイン のOAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office UX 設計パターン コードの追加](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
