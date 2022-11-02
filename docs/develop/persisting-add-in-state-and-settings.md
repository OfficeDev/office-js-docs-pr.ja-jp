---
title: アドインの状態と設定を保持する
description: ブラウザー コントロールのステートレス環境で実行されている Office アドイン Web アプリケーションでデータを保持する方法について説明します。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e2018e5ecf419744257cdceac31b8b1688fa65ff
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810009"
---
# <a name="persist-add-in-state-and-settings"></a>アドインの状態と設定を保持する

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.
To do that, you can:

- データを格納する Office JavaScript API のメンバーを次のように使用します。
  - アドインの種類応じた場所に保存されるプロパティ バッグ内の名前と値の組。
  - ドキュメント内に保存されるカスタム XML。

- 基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) または [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) を使用します。
    > [!NOTE]
    > ブラウザーまたはユーザーのブラウザー設定によっては、ブラウザー ベースのストレージ手法がブロックされる場合があります。 [「Web Storage API の使用](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API)」に記載されているように、可用性をテストする必要があります。

この記事では、Office JavaScript API を使用してアドインの状態を現在のドキュメントに保持する方法について説明します。 開いているドキュメント間でユーザー設定を追跡するなど、ドキュメント間で状態を保持する必要がある場合は、別の方法を使用する必要があります。 たとえば、 [SSO を](use-sso-to-get-office-signed-in-user-token.md) 使用してユーザー ID を取得し、ユーザー ID とその設定をオンライン データベースに保存できます。

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>Office JavaScript API を使用してアドインの状態と設定を保持する

Office JavaScript API には、次の表に示すように、セッション間でアドインの状態を保存するための [Settings](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、 [CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトが用意されています。 すべてのケースで、保存された設定値は、それを作成したアドインの [Id](/javascript/api/manifest/id) にのみ関連付けられます。

|オブジェクト|アドインの種類のサポート|保存場所|Office アプリケーションのサポート|
|:-----|:-----|:-----|:-----|
|[設定](/javascript/api/office/office.settings)|-コンテンツ<br>- 作業ウィンドウ|アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。 コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|-単語<br>-Excel<br>-Powerpoint<br/><br/> **メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。 ただし、Project (およびその他の Office クライアント アプリケーション) で実行されているアドインでは、ブラウザー Cookie や Web ストレージなどの手法を使用できます。 こうした技術の詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|mail|アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。 これらの設定はユーザーのサーバー メールボックスに格納されるため、ユーザーと一緒に "ローミング" でき、サポートされている Office クライアント アプリケーションまたはそのユーザーのメールボックスにアクセスするブラウザーのコンテキストで実行されているアドインで使用できます。<br/><br/> Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|mail|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|作業ウィンドウ|The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|- Word (Office JavaScript Common API を使用)<br>- Excel (アプリケーション固有の Excel JavaScript API を使用)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>実行時のメモリ内での設定データの管理

> [!NOTE]
> この後の 2 つのセクションでは、Office 共通 JavaScript API のコンテキストでの設定について説明します。 アプリケーション固有の Excel JavaScript API では、カスタム設定へのアクセスも提供されます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の SettingCollection](/javascript/api/excel/excel.settingcollection) を参照してください。

内部的には、または `RoamingSettings` オブジェクトでアクセスされるプロパティ バッグ内の`Settings``CustomProperties`データは、名前と値のペアを含むシリアル化された JavaScript Object Notation (JSON) オブジェクトとして格納されます。 各値の名前 (キー) は である`string`必要があり、格納される値には JavaScript`string`、、`number``date`または `object`を指定できますが、**関数** は使用できません。

この例はプロパティ バッグの構造を示し、3 つの定義された **string** 値 (`firstName`、`location`、`defaultView` という名前) が含まれます。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

設定プロパティ バッグは、前のアドイン セッション中に保存された後、アドインが初期化されるとき、またはその後はいつでも、アドインの現行セッション中は読み込むことができます。 セッション中、設定は、作成する設定の種類 (**設定**、**CustomProperties**、または **RoamingSettings**) に対応する オブジェクトの メソッドを使用して`get``set``remove`、完全にメモリ内で管理されます。

> [!IMPORTANT]
> アドインの現在のセッション中に行われた追加、更新、または削除をストレージの場所に保持するには、その種類の設定を処理するために使用する対応するオブジェクトのメソッドを呼び出す `saveAsync` 必要があります。 、`set`、および `remove` メソッドは`get`、設定プロパティ バッグのメモリ内コピーでのみ動作します。 を呼び出 `saveAsync`さずにアドインを閉じると、そのセッション中に設定に加えられた変更は失われます。

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法

Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、 [Settings](/javascript/api/office/office.settings) オブジェクトとそのメソッドを使用します。 オブジェクトの `Settings` メソッドを使用して作成されたプロパティ バッグは、オブジェクトを作成したコンテンツアドインまたは作業ウィンドウ アドインのインスタンスでのみ使用でき、保存されているドキュメントからのみ使用できます。

オブジェクトは `Settings` [Document](/javascript/api/office/office.document) オブジェクトの一部として自動的に読み込まれ、作業ウィンドウまたはコンテンツ アドインがアクティブになったときに使用できます。 オブジェクトが `Document` インスタンス化された後、オブジェクトの `Settings` [settings](/javascript/api/office/office.document#office-office-document-settings-member) プロパティを使用して オブジェクトに `Document` アクセスできます。 セッションの有効期間中は、 メソッド、、および `Settings.remove` メソッドを使用`Settings.set``Settings.get`するだけで、プロパティ バッグのメモリ内コピーから永続化された設定とアドインの状態を読み取り、書き込み、または削除できます。

set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) メソッドを呼び出す必要があります。

### <a name="creating-or-updating-a-setting-value"></a>設定値の作成または更新

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。 メソッドを `Settings.saveAsync` 使用して、新しい設定または更新された設定をドキュメントに保持します。

### <a name="getting-the-value-of-a-setting"></a>設定値の取得

次の例では、 [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。 メソッドの唯一の `get` パラメーターは、設定の大文字と小文字を区別する _名前_ です。

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 メソッドは `get` 、渡された設定 _名_ に対して以前に保存された値を返します。 設定が存在しない場合、メソッドは **null** を返します。

### <a name="removing-a-setting"></a>設定の削除

次の例では、 [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。 メソッドの唯一の `remove` パラメーターは、設定の大文字と小文字を区別する _名前_ です。

```js
Office.context.document.settings.remove('themeColor');
```

該当する設定が存在しない場合は何も起きません。 メソッドを `Settings.saveAsync` 使用して、ドキュメントからの設定の削除を保持します。

### <a name="saving-your-settings"></a>設定の保存

現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。 メソッドの唯一の `saveAsync` パラメーターは _callback_ です。これは、1 つのパラメーターを持つコールバック関数です。

```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

_コールバック_ パラメーターとしてメソッドに`saveAsync`渡される匿名関数は、操作の完了時に実行されます。 コールバックの _asyncResult_ パラメーターは、操作の状態を `AsyncResult` 含むオブジェクトへのアクセスを提供します。 この例では、 関数によって プロパティが `AsyncResult.status` チェックされ、保存操作が成功したか失敗したかが確認され、その結果がアドインのページに表示されます。

## <a name="how-to-save-custom-xml-to-the-document"></a>ドキュメントにカスタム XML を保存する方法

> [!NOTE]
> このセクションでは、Word でサポートされている Office 共通 JavaScript API のコンテキストでのカスタム XML 部分について説明します。 アプリケーション固有の Excel JavaScript API では、カスタム XML パーツへのアクセスも提供されます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の CustomXmlPart](/javascript/api/excel/excel.customxmlpart) を参照してください。

ドキュメント設定のサイズ制限を超える情報や構造化文字を含む情報を格納する必要がある場合は、追加のストレージ オプションがあります。 Word および Excel の作業ウィンドウ アドインには、カスタムの XML マークアップを保持できます (Excel については、このセクションの冒頭にあるノートを参照してください)。 Word の場合は、[CustomXmlPart](/javascript/api/office/office.customxmlpart) とそのメソッドを使用します (繰り返しになりますが、Excel の場合は上記のノートを参照してください)。 次のコードでは、カスタム XML パーツを作成して、その ID とコンテンツをページの div に表示します。 XML 文字列には `xmlns` 属性が必ず存在する点に注意してください。

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

カスタム XML 部分を取得するには、[getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) メソッドを使用しますが、ID は XML 部分の作成時に生成された GUID になるため、コードの作成時に ID の内容を知ることはできません。 そのため、XML 部分を作成したら、その XML 部分の ID を設定としてすぐに保存して、覚えやすいキーを割り当てることがベスト プラクティスになります。 次のメソッドは、この方法を示してます  (ただし、カスタム設定を使用する場合の詳細とベスト プラクティスについては、この記事の前のセクションを参照してください)。

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

次のコードは、最初に設定から ID を取得することで、XML 部分を取得する方法を示しています。

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Outlook アドインに設定を保存する方法

Outlook アドインに設定を保存する方法の詳細については、「Outlook アドインの [状態と設定を管理する](../outlook/manage-state-and-settings-outlook.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Outlook アドイン](../outlook/outlook-add-ins-overview.md)
- [Outlook アドインの状態と設定を管理する](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
