---
title: Office アドインで Office ダイアログ API を使用する
description: Office アドインでダイアログボックスを作成する方法の基本事項について説明します。
ms.date: 06/10/2020
localization_priority: Normal
ms.openlocfilehash: 749fd6041c2ef60a4d766e865e25d53e97298d01
ms.sourcegitcommit: 449a728118db88dea22a44f83728d21604d6ee8c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2020
ms.locfileid: "44719071"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Office アドインで Office ダイアログ API を使用する

[Office ダイアログ API](/javascript/api/office/office.ui) を使用して、Office アドインでダイアログ ボックスを開くことができます。 この記事では、Office アドインでダイアログ API を使用するためのガイダンスを提供します。

> [!NOTE]
> ダイアログ API の現在のサポート状態に関する詳細は、「[ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)」を参照してください。現在、ダイアログ API は Word、Excel、PowerPoint、および Outlook でサポートされています。

ダイアログ API の主要なシナリオは、Google や Facebook、Microsoft Graph などのリソースで認証を有効にすることです。 詳細については、この記事をよく読んだ*後*で「[Office Dialog API を使用して認証する](auth-with-office-dialog-api.md)」を参照してください。

作業ウィンドウ アドイン、コンテンツ アドイン、[アドイン コマンド](../design/add-in-commands.md)からダイアログ ボックスを開いて、次の操作を実行することを検討してください。

- 作業ウィンドウに直接開くことができないサインイン ページを表示する。
- アドインでの作業用に画面領域を広げる (あるいは全画面表示)。
- ビデオが作業ウィンドウに限定されている場合に、小さすぎるビデオをホストする。

> [!NOTE]
> UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 例については、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) のサンプルを参照してください。

次の画像は、ダイアログ ボックスの例を示します。

![アドイン コマンド](../images/auth-o-dialog-open.png)

ダイアログ ボックスが常に画面の中央に開くことに注意してください。 ユーザーはダイアログ ボックスの移動とサイズ変更ができます。 ウィンドウは*モードレス*です。ホスト Office アプリケーションのドキュメントの操作と、作業ウィンドウのページ (存在する場合) の操作の両方を続行できます。

## <a name="open-a-dialog-box-from-a-host-page"></a>ホスト ページからダイアログ ボックスを開く

Office JavaScript API には、[Dialog](/javascript/api/office/office.dialog) オブジェクトと [Office.context.ui 名前空間](/javascript/api/office/office.ui)の 2 つの関数が含まれます。

ダイアログ ボックスを開くには、コード (通常は作業ウィンドウ内のページ) で [displayDialogAsync](/javascript/api/office/office.ui) メソッドを呼び出して、開くリソースの URL を渡します。 このメソッドを呼び出すページは、「ホスト ページ」と呼ばれます。 たとえば、作業ウィンドウの index.html にあるスクリプトでこのメソッドを呼び出した場合は、index.html がメソッドが開いたダイアログ ボックスのホスト ページです。

ダイアログ ボックスで開かれるリソースは通常ページですが、MVC アプリケーションのコントローラー メソッド、ルート、Web サービス メソッド、またはその他のリソースの場合もあります。 この記事では、'ページ' または 'Web サイト' とは、ダイアログ ボックス内のリソースを意味します。 次のコードは簡単な例を示しています。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - URL には HTTP**S** プロトコルを使用します。 これは、読み込まれる最初のページだけでなく、ダイアログ ボックスに読み込まれるすべてのページに対して必須です。
> - ダイアログ ボックスのドメインはホスト ページのドメインと同じです。ホスト ページは、作業ウィンドウ内のページまたはアドイン コマンドの[関数ファイル](../reference/manifest/functionfile.md)にすることができます。 ページ、コントローラーのメソッド、または `displayDialogAsync` メソッドに渡されるその他のリソースは、ホスト ページと同じドメインにある必要があります。

> [!IMPORTANT]
> ダイアログ ボックスで開くホスト ページとリソースのフル ドメインは、同じである必要があります。 `displayDialogAsync` にアドインのドメインのサブドメインを渡そうとすると、正常に動作しません。 サブドメインを含む、フル ドメインが一致している必要があります。

最初のページ (または他のリソース) が読み込まれると、ユーザーはリンクまたは他の UI を使用して HTTPS を使用する任意の Web サイト (または他のリソース) に移動できます。 また、すぐに別のサイトにリダイレクトするように最初のページを設計することもできます。

既定では、ダイアログ ボックスのサイズはデバイス画面の高さと幅の 80% ですが、次の例に示すように、メソッドに構成オブジェクトを渡すことによってさまざまな割合を設定できます。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

これを実行するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

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
> どの時点であれ、iframe で開けないページにダイアログ ボックスがリダイレクトされることになる場合は、`displayInIframe: true` を使用すべきでは**ありません**。 たとえば、Google や Microsoft アカウントなどの多くの一般的な Web サービスのサインイン ページは iframe で開くことができません。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>ダイアログ ボックスからホスト ページに情報を送信する

ダイアログ ボックスは、以下の場合を除いて、作業ウィンドウのホスト ページと通信できません。

- ダイアログ ボックスの現在のページがホスト ページと同じドメインにある。
- Office JavaScript API ライブラリがページにロードされます。(Office JavaScript API ライブラリを使用するページと同様に、ページのスクリプトはメソッドをプロパティに割り当てる必要がありますが、空のメソッドにする `Office.initialize` こともできます。詳細については、「 [Office アドインを初期化する](initialize-add-in.md)」を参照してください)。

ダイアログ ボックスのコードは、[messageParent](/javascript/api/office/office.ui#messageparent-message-) 関数を使用して、ブール値または文字列メッセージのいずれかをホスト ページに送信します。 文字列には、単語、文、XML BLOB、文字列に変換された JSON、文字列にシリアル化できるすべてのものを指定できます。 例を次に示します。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - `messageParent` 関数を呼び出せるのは、ホスト ページと同じドメイン (プロトコルとポートを含む) を持つページ上のみです。
> - この `messageParent` 関数は、ダイアログ*only*ボックスで呼び出すことができる2つの Office JS api のうちの1つです。 
> - ダイアログボックスで呼び出すことができるその他の JS API は、 `Office.context.requirements.isSetSupported` です。 詳細は、「[Office のホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してださい。 ただし、ダイアログボックスでは、この API は Outlook 2016 1 での購入時 (つまり、MSI バージョン) ではサポートされていません。

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
> - Office は `arg` オブジェクトをハンドラーに渡します。 その `message` プロパティは、ダイアログ ボックスの `messageParent` の呼び出しで送信されるブール値または文字列です。 この例では、Microsoft アカウントまたは Google などのサービスからのユーザーのプロファイルの文字列に変換された表記です。このため、`JSON.parse` を含むオブジェクトに逆シリアル化されます。
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

ホスト ページがダイアログ ボックスに情報を渡す必要がある場合もあります。これは主に 2 つの方法で実行することができます。

- `displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。
- ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。 2 つのウィンドウは共通のセッション ストレージを共有しませんが、ポート番号 (存在する場合) を含む*ドメインが同じである場合*は、共通の[ローカル ストレージ](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。\*

> [!NOTE]
> \* トークン処理の戦略に影響を与えるバグがあります。 Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。

### <a name="use-local-storage"></a>ローカル ストレージの使用

ローカル ストレージを使用するには、次の例に示すように、`displayDialogAsync` 呼び出しの前に、コードはホスト ページで `window.localStorage` オブジェクトの `setItem` メソッドを呼び出します。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

ダイアログ ボックス内のコードは、次の例に示すように、必要に応じて項目を読み取ります。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a>クエリ パラメーターの使用

次の例では、クエリ パラメーターを使用してデータを渡す方法を示します。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。

ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。

> [!IMPORTANT]
> Office は、`displayDialogAsync` に渡される URL に `_host_info` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。ダイアログ ボックスが移動する先の後続の URL には追加されません)。Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。ダイアログ ボックスのセッション ストレージには、同じ値が追加されます。この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

> [!NOTE]
> これで、 `messageChild` `messageParent` 上記の api がダイアログからメッセージを送信するのと同じように、親ページがダイアログにメッセージを送信するために使用できる api がプレビューされました。 詳細については、「[ホストページからダイアログボックスにデータとメッセージを渡す](parent-to-dialog.md)」を参照してください。 試してみることをお勧めしますが、運用アドインでは、このセクションで説明する手法を使用することをお勧めします。

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
