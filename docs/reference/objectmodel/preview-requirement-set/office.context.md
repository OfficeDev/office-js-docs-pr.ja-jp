---
title: Office コンテキスト-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629189"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="properties"></a>プロパティ

| プロパティ | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|
| [contentLanguage](#contentlanguage-string) | 作成<br>読み取り | String | 1.0 |
| [ダン](#diagnostics-contextinformation) | 作成<br>読み取り | [ContextInformation](/javascript/api/office/office.contextinformation) | 1.0 |
| [displayLanguage](#displaylanguage-string) | 作成<br>読み取り | String | 1.0 |
| [主催](#host-hosttype) | 作成<br>読み取り | [HostType](/javascript/api/office/office.hosttype) | 1.0 |
| [officeTheme](#officetheme-officetheme) | 作成<br>読み取り | [OfficeTheme](/javascript/api/office/office.officetheme) | プレビュー |
| [プラットフォーム](#platform-platformtype) | 作成<br>読み取り | [PlatformType](/javascript/api/office/office.platformtype) | 1.0 |
| [要件](#requirements-requirementsetsupport) | 作成<br>読み取り | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport) | 1.0 |
| [roamingSettings](#roamingsettings-roamingsettings) | 作成<br>読み取り | [RoamingSettings](/javascript/api/outlook/office.roamingsettings) | 1.0 |
| [UI](#ui-ui) | 作成<br>読み取り | [UI](/javascript/api/office/office.ui) | 1.0 |

### <a name="namespaces"></a>名前空間

[auth](/javascript/api/office/office.auth):[シングルサインオン (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token)のサポートを提供します。

[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。

## <a name="property-details"></a>プロパティの詳細

#### <a name="contentlanguage-string"></a>contentLanguage: String

アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。

この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
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

<br>

---
---

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a>診断: [Contextinformation](/javascript/api/office/office.contextinformation)

アドインが実行されている環境に関する情報を取得します。

##### <a name="type"></a>型

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: String

Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。

`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

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

<br>

---
---

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a>ホスト: [Hosttype](/javascript/api/office/office.hosttype)

アドインが実行されている Office アプリケーションホストを取得します。

##### <a name="type"></a>型

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a>officeTheme: [Officetheme](/javascript/api/office/office.officetheme)

Office テーマの色のプロパティにアクセスできるようにします。

> [!NOTE]
> このメンバーは、Windows の Outlook でのみサポートされています。

Office テーマの色を使用すると、アドインの配色を、[**ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office ホストアプリケーションで適用されます。 Using Office theme colors is appropriate for mail and task pane add-ins.

##### <a name="type"></a>型

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`bodyBackgroundColor`| String|Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。|
|`bodyForegroundColor`| String|Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。|
|`controlBackgroundColor`| String|Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。|
|`controlForegroundColor`| String|Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。|

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| プレビュー|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

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

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a>プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)

アドインが実行されているプラットフォームを提供します。

##### <a name="type"></a>型

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a>要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。

##### <a name="type"></a>型

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)

ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。

`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。

##### <a name="type"></a>型

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a>ui: [ui](/javascript/api/office/office.ui)

Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。

##### <a name="type"></a>型

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|
