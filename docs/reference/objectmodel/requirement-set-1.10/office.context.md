---
title: Office.context - 要件セット 1.10
description: Office。メールボックス API 要件セット 1.10 をOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937809"
---
# <a name="context-mailbox-requirement-set-110"></a>context (メールボックス要件セット 1.10)

### <a name="officecontext"></a>[Office](office.md).context

Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。 この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true)。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|:---:|
| [auth](#auth-auth) | 作成<br>読み取り | [Auth](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [診断](#diagnostics-contextinformation) | 作成<br>読み取り | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | 作成<br>読み取り | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | 作成<br>読み取り | [メールボックス](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [platform](#platform-platformtype) | 作成<br>読み取り | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [要件](#requirements-requirementsetsupport) | 作成<br>読み取り | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | 作成<br>読み取り | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [UI](#ui-ui) | 作成<br>読み取り | [UI](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>プロパティの詳細

#### <a name="auth-auth"></a>auth: [Auth](/javascript/api/office/office.auth)

シングル[サインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするには、Office アプリケーションがアドインの Web アプリケーションへのアクセス トークンを取得できるメソッドを提供します。 これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。

##### <a name="type"></a>型

*   [Auth](/javascript/api/office/office.auth)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.10|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a>contentLanguage: String

アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。

この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

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

#### <a name="diagnostics-contextinformation"></a>診断: [ContextInformation](/javascript/api/office/office.contextinformation)

アドインが実行されている環境に関する情報を取得します。

##### <a name="type"></a>型

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: String

ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。

この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

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

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype)

アドインをOfficeしているアプリケーションを取得します。

> [!NOTE]
> または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してホストを取得できます。

##### <a name="type"></a>型

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)

アドインが実行されているプラットフォームを提供します。

> [!NOTE]
> または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。

##### <a name="type"></a>型

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。

##### <a name="type"></a>型

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)

ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。

このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`

##### <a name="type"></a>型

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="ui-ui"></a>ui: [UI](/javascript/api/office/office.ui)

ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。

##### <a name="type"></a>型

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|
