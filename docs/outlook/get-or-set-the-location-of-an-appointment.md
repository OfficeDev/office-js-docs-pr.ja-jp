---
title: アドインで予定の場所を取得または設定する
description: Outlook アドインで予定の場所を取得または設定する方法について説明します。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: bf03e0e470bb5aea811c09bb7b88cc5a915a7a13
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541240"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに場所を取得または設定する

Office JavaScript API には、ユーザーが作成している予定の場所を管理するためのプロパティとメソッドが用意されています。 現在、予定の場所を提供する 2 つのプロパティがあります。

- [item.location: 場所](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)を取得および設定できる基本的な API。
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): 場所の取得と設定を可能にし、場所の [種類](/javascript/api/outlook/office.mailboxenums.locationtype)の指定を含む拡張 API。 種類は、 `LocationType.Custom` 場所を設定する場合に使用します `item.location`。

次の表に、場所 API と、使用可能なモード (つまり、作成または読み取り) を示します。

| API | 適用可能な予定モード |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | 出席者/読み取り |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | 開催者/Compose |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | 開催者/Compose |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | オーガナイザー/Compose,<br>出席者/読み取り |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | 開催者/Compose |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | 開催者/Compose |

アドインの作成にのみ使用できるメソッドを使用するには、アドインのオーガナイザー/作成モードをアクティブ化するようにアドイン XML マニフェストを構成します。 詳細については、「 [作成フォーム用の Outlook アドインの作成](compose-scenario.md) 」を参照してください。 Office アドイン [の Teams マニフェスト (プレビュー) を](../develop/json-manifest-overview.md)使用するアドインでは、アクティブ化ルールはサポートされていません。

## <a name="use-the-enhancedlocation-api"></a>API を使用する`enhancedLocation`

API を `enhancedLocation` 使用して、予定の場所を取得および設定できます。 場所フィールドは複数の場所をサポートし、場所ごとに表示名、種類、会議室の電子メール アドレス (該当する場合) を設定できます。 サポートされている場所の種類については、 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) を参照してください。

### <a name="add-location"></a>場所を追加する

次の例は、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) で [addAsync を](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1))呼び出して場所を追加する方法を示しています。

```js
let item;
const locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>場所を取得する

次の例は、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member) で [getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) を呼び出して場所を取得する方法を示しています。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a>場所を削除する

次の例は、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) で [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) を呼び出して場所を削除する方法を示しています。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>API を使用する`location`

API を `location` 使用して、予定の場所を取得および設定できます。

### <a name="get-the-location"></a>場所を取得する

ここでは、ユーザーが新規作成している予定の配置場所を取得し、それを表示するコード サンプルを示します。

使用 `item.location.getAsync`するには、非同期呼び出しの状態と結果を確認するコールバック関数を指定します。 省略可能なパラメーターを使用して、コールバック関数に必要な引数を `asyncContext` 指定できます。 コールバックの出力パラメーター `asyncResult` を使用して、状態、結果、エラーを取得できます。 非同期コールが成功した場合、[AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) プロパティを使用して、配置場所を文字列として取得することができます。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a>場所を設定する

ここでは、ユーザーが新規作成している予定の配置場所を設定するコード サンプルを示します。

`item.location.setAsync` を使用するには、data パラメーターに最大 255 文字までの文字列を指定します。 必要に応じて、コールバック関数と、パラメーター内のコールバック関数の任意の引数を `asyncContext` 指定できます。 コールバックの出力パラメーターの `asyncResult` 状態、結果、エラー メッセージを確認する必要があります。 非同期呼び出しが成功した場合、`setAsync` はそのアイテムの既存の配置場所を上書きし、指定した配置場所をプレーンテキストとして挿入します。

> [!NOTE]
> 区切り記号としてセミコロンを使用して複数の場所を設定できます (例: '会議室 A;会議室 B')

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a>関連項目

- [初めての Outlook アドインを作成する](../quickstarts/outlook-quickstart.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
