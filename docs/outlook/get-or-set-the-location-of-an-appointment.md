---
title: アドインで予定の場所を取得または設定する
description: Outlook アドインで予定の場所を取得または設定する方法について説明します。
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: d88e2494592d9b261945ecdaf0ca27ae79c73ba8
ms.sourcegitcommit: cae583433e489a3b71418ea270a90db72ad1e838
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892365"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに場所を取得または設定する

Office JavaScript API には、ユーザーが作成している予定の場所を管理するためのプロパティとメソッドが用意されています。 現在、予定の場所を提供するプロパティは 2 つあります。

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): 場所を取得して設定できる基本的な API。
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): 場所の取得と設定を可能にし、場所の [種類](/javascript/api/outlook/office.mailboxenums.locationtype)の指定を含む拡張 API。 型は、 を使用して`item.location`場所を設定する場合です`LocationType.Custom`。

次の表に、場所 API と、使用可能なモード (新規作成または読み取り) を示します。

| API | 適用可能な予定モード |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | Attendee/Read |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | Organizer/Compose |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | Organizer/Compose |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | Organizer/Compose,<br>Attendee/Read |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | Organizer/Compose |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | Organizer/Compose |

アドインの作成にのみ使用できるメソッドを使用するには、アドイン XML マニフェストを構成して、オーガナイザー/新規作成モードでアドインをアクティブにします。 詳細については、「 [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md) 」を参照してください。 アクティブ化ルールは、Office アドイン用 [の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用するアドインではサポートされていません。

## <a name="use-the-enhancedlocation-api"></a>API を使用する`enhancedLocation`

API を `enhancedLocation` 使用して、予定の場所を取得および設定できます。 場所フィールドは複数の場所をサポートしており、場所ごとに表示名、種類、会議室のメール アドレス (該当する場合) を設定できます。 サポートされている場所の種類については、「 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) 」を参照してください。

### <a name="add-location"></a>場所の追加

次の例では、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) で [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) を呼び出して場所を追加する方法を示します。

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

> [!NOTE]
> 予定の場所として追加された[個人用連絡先グループ](https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023)は、[enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) メソッドによって返されません。

### <a name="remove-location"></a>場所を削除する

次の例では、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) で [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) を呼び出して場所を削除する方法を示します。

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

を使用 `item.location.getAsync`するには、非同期呼び出しの状態と結果をチェックするコールバック関数を指定します。 省略可能なパラメーターを使用して、コールバック関数に必要な引数を `asyncContext` 指定できます。 コールバックの出力パラメーター `asyncResult` を使用して、状態、結果、およびエラーを取得できます。 非同期コールが成功した場合、[AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) プロパティを使用して、配置場所を文字列として取得することができます。

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

`item.location.setAsync` を使用するには、data パラメーターに最大 255 文字までの文字列を指定します。 必要に応じて、 パラメーターにコールバック関数とコールバック関数の任意の引数を `asyncContext` 指定できます。 コールバックの出力パラメーターで `asyncResult` 、状態、結果、およびエラー メッセージを確認する必要があります。 非同期呼び出しが成功した場合、`setAsync` はそのアイテムの既存の配置場所を上書きし、指定した配置場所をプレーンテキストとして挿入します。

> [!NOTE]
> 区切り記号としてセミコロンを使用して複数の場所を設定できます (例: '会議室 A;会議室 B')。

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

- [最初の Outlook アドインを作成する](../quickstarts/outlook-quickstart.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
