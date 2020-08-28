---
title: Outlook アドインの状態と設定を管理する
description: Outlook アドインのアドインの状態と設定を保存する方法について説明します。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: b11468cdcd776a2d93bb4a435e5c878ca9ab42a3
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293843"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Outlook アドインの状態と設定を管理する

> [!NOTE]
> この記事を読む前に、このドキュメントの「**コア概念**」セクションの「[アドインの状態と設定を保持](../develop/persisting-add-in-state-and-settings.md)する」を参照してください。

Outlook アドインの場合、Office JavaScript API は、次の表に示すように、セッション間でアドインの状態を保存するための [RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトと [CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトを提供します。 すべてのケースで、保存された設定値は、それを作成したアドインの [Id](../reference/manifest/id.md) にのみ関連付けられます。

|**オブジェクト**|**ストレージの場所**|
|:-----|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。 これらの設定はユーザーのサーバーメールボックスに格納されるため、ユーザーとの "ローミング" が可能であり、そのユーザーのメールボックスにアクセスするサポートされているクライアントのコンテキストで実行されている場合は、アドインで使用できます。<br/><br/> Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Outlook アドインでユーザーのメールボックスに設定をローミング設定として保存する方法

Outlook アドインは、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトを使用して、ユーザーのメールボックスに固有の、アドインの状態および設定のデータを保存できます。 このデータには、アドインを実行しているユーザーではなく、Outlook アドインのみがアクセスできます。 データはユーザーの Exchange Server メールボックスに格納されます。データには、ユーザーが自分のアカウントにログインして Outlook アドインを実行したときにアクセスできるようになります。

### <a name="loading-roaming-settings"></a>ローミング設定の読み込み

通常、Outlook アドインでは、 [Office.initialize](/javascript/api/office) イベント ハンドラーでローミング設定を読み込みます。次の JavaScript のコード例は、既存のローミング設定を読み込む方法を示しています。

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a>ローミング設定の作成または割り当て

前の例に続けて、次の  `setAppSetting` 関数では、 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) メソッドを使用して、 `cookie` という名前の設定項目に今日の日付を設定、または今日の日付で更新する方法を示しています。次に、 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) メソッドを使用して Exchange Server にすべてのローミング設定を保存し直しています。

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

**saveAsync** メソッドは、ローミング設定を非同期で保存し、オプションのコールバック関数を受け取ります。 このコード例では、`saveMyAppSettingsCallback` という名前のコールバック関数を **saveAsync** メソッドに渡します。 非同期呼び出しが返されると、`saveMyAppSettingsCallback` 関数の _asyncResult_ パラメーターが [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトにアクセスします。このオブジェクトを使用すると、**AsyncResult.status** プロパティで操作の成功または失敗を判定することができます。

### <a name="removing-a-roaming-setting"></a>ローミング設定の削除

また、次の  `removeAppSetting` 関数は、前の例をさらに拡張するものです。この例では、 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に保存し直す方法を示しています。

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Outlook アドインでアイテムごとに設定をカスタムプロパティとして保存する方法

カスタム プロパティを使用すると、Outlook アドインは処理しているアイテムに関する情報を保存できます。たとえば、Outlook アドインを使用して、メッセージ内の会議の提案から予定を作成する場合は、カスタム プロパティを使用して、会議が作成されたという事実を保存できます。これにより、メッセージを再び開いたときに、Outlook アドインが再び予定の作成を行うことはありません。

メッセージ、予定、または会議出席依頼の特定のアイテムに対してカスタム プロパティを使用するには、その前に、 [Item](/javascript/api/outlook/office.mailbox) オブジェクトの **loadCustomPropertiesAsync** メソッドを呼び出して、プロパティをメモリに読み込む必要があります。現在のアイテムに対してカスタム プロパティが既に設定されている場合は、この時点で Exchange サーバーから読み込まれます。プロパティを読み込んだ後、 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) オブジェクトの [set](/javascript/api/outlook/office.roamingsettings) メソッドおよび **get** メソッドを使用して、メモリ内のプロパティの追加、更新、および取得を実行できます。アイテムのカスタム プロパティに対して行った変更を保存するには、 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) メソッドを使用して、アイテムに加えた変更を Exchange サーバー上で保持する必要があります。

### <a name="custom-properties-example"></a>カスタム プロパティの例

以下の例では、カスタム プロパティを使用する Outlook アドインの一連の関数を、簡略化して示しています。この例を出発点として、カスタム プロパティを使用する Outlook アドインを作成できます。 

これらの関数を使用する Outlook アドインは、次の例に示すように、`_customProps` 変数で **get** メソッドを呼び出すことによって、任意のカスタム プロパティを取得します。

```js
var property = _customProps.get("propertyName");
```

以下の例には、次の関数が含まれています。

|**関数名**|**説明**|
|:-----|:-----|
| `Office.initialize`|アドインを初期化し、Exchange サーバーから現在のアイテムのカスタム プロパティを読み込みます。|
| `customPropsCallback`|Exchange サーバーから返されるカスタム プロパティを取得し、後で使用できるように保存します。|
| `updateProperty`|特定のプロパティを設定または更新し、その変更を Exchange サーバーに保存します。|
| `removeProperty`|特定のプロパティを削除し、その削除を Exchange サーバーに保存します。|
| `saveCallback`|`updateProperty` 関数および `removeProperty` 関数内で **saveAsync** メソッドを呼び出すためのコールバック。|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

## <a name="see-also"></a>関連項目

- [アドインの状態および設定を保持する](../develop/persisting-add-in-state-and-settings.md)
- [Office アドインを初期化する](../develop/initialize-add-in.md)
