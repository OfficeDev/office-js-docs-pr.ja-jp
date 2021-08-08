---
title: アドインの状態および設定を保持する
description: ブラウザー コントロールのステートレスOfficeで実行されているアドイン Web アプリケーションにデータを保持する方法について説明します。
ms.date: 03/23/2021
localization_priority: Normal
ms.openlocfilehash: 5b2a7d59340a0126fe270334c560ba5fbf1f1739ef6f119a2cef754a06321699
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080283"
---
# <a name="persisting-add-in-state-and-settings"></a>アドインの状態および設定を保持する

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office アドインは、基本的にブラウザー コントロールのステートレス環境で動作する Web アプリケーションです。したがって、アドインでは、そのアドインを使用するセッション間で特定の操作または機能を継続して維持するためのデータを保持することが必要な場合があります。たとえば、アドインには、ユーザーの優先ビューや既定の場所など、アドインで保存しておき、アドインが次回初期化されたときにリロードする必要があるカスタム設定または他の値が含まれる場合があります。その場合は、次のようにします。

- データを次のように格納Office JavaScript API のメンバーを使用します。
  - アドインの種類応じた場所に保存されるプロパティ バッグ内の名前と値の組。
  - ドキュメント内に保存されるカスタム XML。

- 基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) または [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) を使用します。
    > [!NOTE]
    > ユーザーは、選択した設定に応じて、ブラウザー ベースのストレージ手法をブロックできます。

この記事では、JavaScript API Officeを使用して現在のドキュメントにアドインの状態を保持する方法について説明します。 開いているドキュメントのユーザー設定の追跡など、ドキュメント間で状態を保持する必要がある場合は、別の方法を使用する必要があります。 たとえば [、SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) を使用してユーザー ID を取得し、ユーザー ID とその設定をオンライン データベースに保存できます。

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>JavaScript API を使用したアドインの状態とOffice保持する

JavaScript API Officeには、設定、RoamingSettings、CustomProperties[](/javascript/api/outlook/office.customproperties)オブジェクトが提供され、次の表に示すとおり、セッション間でアドインの状態を保存できます。 [](/javascript/api/office/office.settings) [](/javascript/api/outlook/office.roamingsettings) すべてのケースで、保存された設定値は、それを作成したアドインの [Id](../reference/manifest/id.md) にのみ関連付けられます。

|**オブジェクト**|**アドインの種類のサポート**|**ストレージの場所**|**Officeサポート**|
|:-----|:-----|:-----|:-----|
|[Settings](/javascript/api/office/office.settings)|コンテンツおよび作業ウィンドウ|アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。 コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。<br/><br/>**重要:****Settings** オブジェクトを使用して、パスワードおよびその他の機密の個人を特定できる情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されているため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。|Word、Excel、または PowerPoint<br/><br/> **メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。 ただし、Project (および他の Office クライアント アプリケーション) で実行されているアドインでは、ブラウザー Cookie や Web ストレージなどの手法を使用できます。 こうした技術の詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。 これらの設定はユーザーのサーバー メールボックスに格納されますので、ユーザーと一緒に "ローミング" し、サポートされている Office クライアント アプリケーションまたはブラウザーがユーザーのメールボックスにアクセスするコンテキストでアドインを実行している場合に使用できます。<br/><br/> Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|作業ウィンドウ|アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。<br/><br/>**重要:** カスタム XML 部分には、パスワードなどの個人情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されるため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。|Word (JavaScript Office API を使用) Excel (JavaScript API のアプリケーション固有のExcel使用)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>実行時のメモリ内での設定データの管理

> [!NOTE]
> この後の 2 つのセクションでは、Office 共通 JavaScript API のコンテキストでの設定について説明します。 JavaScript API のアプリケーション固有Excelカスタム設定へのアクセスも提供します。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の SettingCollection](/javascript/api/excel/excel.settingcollection) を参照してください。

内部的には、 、 オブジェクトでアクセスされるプロパティ バッグ内のデータは、名前と値のペアを含むシリアル化 `Settings` `CustomProperties` された JavaScript オブジェクト表記 (JSON) オブジェクトとして `RoamingSettings` 格納されます。 各値の名前 (キー) は、 で、格納されている値は JavaScript 、関数である `string` `string` `number` `date` `object` **必要があります**。

この例はプロパティ バッグの構造を示し、3 つの定義された **string** 値 (`firstName`、`location`、`defaultView` という名前) が含まれます。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

設定プロパティ バッグは、前のアドイン セッション中に保存された後、アドインが初期化されるとき、またはその後はいつでも、アドインの現行セッション中は読み込むことができます。 セッション中、設定は、作成する設定の種類 `get` `set` `remove` **(設定、CustomProperties、****または RoamingSettings)** に対応するオブジェクトのメソッドを使用して、メモリ内で完全に管理されます。

> [!IMPORTANT]
> アドインの現在のセッション中に行われた追加、更新、または削除を保存場所に保持するには、その種類の設定を処理するために使用する対応するオブジェクトのメソッドを呼び出す必要があります `saveAsync` 。 、およびメソッドは、settings プロパティ バッグのメモリ内 `get` `set` `remove` コピーでのみ動作します。 アドインを呼び出さずに閉じた場合、そのセッション中に設定に加えた変更 `saveAsync` は失われます。

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法

Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、 [Settings](/javascript/api/office/office.settings) オブジェクトとそのメソッドを使用します。 オブジェクトのメソッドで作成されたプロパティ バッグは、そのオブジェクトを作成したコンテンツ アドインまたは作業ウィンドウ アドインのインスタンスでのみ使用できます。保存されているドキュメントからのみ使用 `Settings` できます。

オブジェクトは Document オブジェクトの一部として自動的に読み込まれ、作業ウィンドウアドインまたはコンテンツ アドインがアクティブ化されると `Settings` 使用できます。 [](/javascript/api/office/office.document) オブジェクトを `Document` インスタンス化した後、オブジェクトの settings プロパティ `Settings` を[](/javascript/api/office/office.document#settings)使用してオブジェクトにアクセス `Document` できます。 セッションの有効期間中に、プロパティ バッグのメモリ内コピーから永続化された設定とアドインの状態を読み取り、書き込み、または削除するには、 およびメソッドを使用できます `Settings.get` `Settings.set` `Settings.remove` 。

set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_) メソッドを呼び出す必要があります。

### <a name="creating-or-updating-a-setting-value"></a>設定値の作成または更新

次のコード例では、[Settings.set](/javascript/api/office/office.settings#set_name__value_) メソッドを使用して `'themeColor'` という名前の設定を作成し、値 `'green'` を指定する方法を説明します。set メソッドの最初のパラメーターは、設定するか作成する設定の _name_ (Id) であり、これは大文字と小文字が区別されます。2 番目のパラメーターは、設定の _value_ です。

```js
Office.context.document.settings.set('themeColor', 'green');
```

 指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。 新しい `Settings.saveAsync` 設定または更新された設定をドキュメントに保持するには、このメソッドを使用します。

### <a name="getting-the-value-of-a-setting"></a>設定値の取得

次の例では、 [Settings.get](/javascript/api/office/office.settings#get_name_) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。 メソッドの唯一の `get` パラメーターは、設定の大文字と小文字を区別する名前です。

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 メソッド `get` は、渡された設定名に対して以前に保存された値を返します。 設定が存在しない場合、メソッドは **null** を返します。

### <a name="removing-a-setting"></a>設定の削除

次の例では、 [Settings.remove](/javascript/api/office/office.settings#remove_name_) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。 メソッドの唯一の `remove` パラメーターは、設定の大文字と小文字を区別する名前です。

```js
Office.context.document.settings.remove('themeColor');
```

該当する設定が存在しない場合は何も起きません。 ドキュメントから `Settings.saveAsync` 設定の削除を保持するには、このメソッドを使用します。

### <a name="saving-your-settings"></a>設定の保存

現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、 [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。 メソッドの唯一の `saveAsync` パラメーターは _、1_ つのパラメーターを持つコールバック関数であるコールバックです。

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

コールバック パラメーターとしてメソッドに渡される匿名関数は、操作 `saveAsync` が完了すると実行されます。  コールバック _の asyncResult_ パラメーターは、操作の状態を含 `AsyncResult` むオブジェクトへのアクセスを提供します。 この例では、この関数はプロパティをチェックして、保存操作が成功したか失敗したのか確認し、その結果をアドインのページ `AsyncResult.status` に表示します。

## <a name="how-to-save-custom-xml-to-the-document"></a>ドキュメントにカスタム XML を保存する方法

> [!NOTE]
> このセクションでは、Word でサポートされている Office 共通 JavaScript API のコンテキストでのカスタム XML 部分について説明します。 JavaScript API のアプリケーションExcelカスタム XML パーツへのアクセスも提供します。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の CustomXmlPart](/javascript/api/excel/excel.customxmlpart) を参照してください。

ドキュメント のサイズ制限を超える情報、または構造化された文字を持つ設定追加の記憶域オプションがあります。 Word および Excel の作業ウィンドウ アドインには、カスタムの XML マークアップを保持できます (Excel については、このセクションの冒頭にあるノートを参照してください)。 Word の場合は、[CustomXmlPart](/javascript/api/office/office.customxmlpart) とそのメソッドを使用します (繰り返しになりますが、Excel の場合は上記のノートを参照してください)。 次のコードでは、カスタム XML パーツを作成して、その ID とコンテンツをページの div に表示します。 XML 文字列には `xmlns` 属性が必ず存在する点に注意してください。

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

カスタム XML 部分を取得するには、[getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_) メソッドを使用しますが、ID は XML 部分の作成時に生成された GUID になるため、コードの作成時に ID の内容を知ることはできません。 そのため、XML 部分を作成したら、その XML 部分の ID を設定としてすぐに保存して、覚えやすいキーを割り当てることがベスト プラクティスになります。 次のメソッドは、この方法を示してます  (ただし、カスタム設定を操作する場合の詳細とベスト プラクティスについては、この記事の前のセクションを参照してください)。

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>アドインに設定をOutlookする方法

カスタム アドインに設定を保存するOutlookについては、「Manage state and settings for an [Outlook」を参照してください](../outlook/manage-state-and-settings-outlook.md)。

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Outlook アドイン](../outlook/outlook-add-ins-overview.md)
- [アドインの状態と設定Outlook管理する](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
