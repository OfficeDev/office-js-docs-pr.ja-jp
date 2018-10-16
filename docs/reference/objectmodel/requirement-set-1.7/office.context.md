
# <a name="context"></a>コンテキスト

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Office.context の名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧では、Outlook アドインで使うインターフェイスのみを記載しています。Office.context の名前空間の完全な一覧については、[ 共有 API の 中のOffice.context リファレンス](/javascript/api/office/office.context) をご覧ください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 型 |
|--------|------|
| [displayLanguage](#displaylanguage-string) | メンバー |
| [officeTheme](#officetheme-object) | メンバー |
| [roamingSettings](#roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings) | メンバー |

### <a name="namespaces"></a>名前空間

[メ―ルボックス](office.context.mailbox.md): Microsoft Outlook と Microsoft Outlook on the Web の Outlook アドイン オブジェクト モデルへアクセスできるようにします。

### <a name="members"></a>メンバー

####  <a name="displaylanguage-string"></a>displayLanguage: 文字列

Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。

値`displayLanguage`は電流を反映する**言語を表示する** Office ホスト アプリケーション内で、**ファイル > オプション > 言語**によって指定された設定

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a>officeTheme: オブジェクト

Office のテーマの色のプロパティにアクセスできるようにします。

> [!NOTE]
> このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。

Office のテーマの色を使うと、**ファイル >Office アカウント >Office テーマ UI**によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Officeの テーマの色は、メール アドインと作業ウィンドウ アドインに適合しています。

##### <a name="type"></a>型:

*   オブジェクト

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`bodyBackgroundColor`| 文字列|Office テーマの本文背景色を 16 進法のカラートリプレットとして取得します。|
|`bodyForegroundColor`| 文字列|Office テーマの本文の前景色を 16 進トリプレットとして取得します。|
|`controlBackgroundColor`| 文字列|Office テーマコントロールの背景色を 16 進法のカラートリプレットとして取得します。|
|`controlForegroundColor`| 文字列|Office テーマの本文コントロール色を 16 進法のカラートリプレットとして取得します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[最小限のメールボックス要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a>roamingSettings:[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)

ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。

`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのためのデータの保存やアクセスができます。そのため、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションからメール アドインを実行してもこのデータを使うことができます。

##### <a name="type"></a>種類:

*   [RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|