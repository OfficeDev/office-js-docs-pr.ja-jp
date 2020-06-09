---
title: アドインで予定の場所を取得または設定する
description: Outlook アドインで予定の場所を取得または設定する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 79cf5ebe029d2b95b1501b6f9066a2c8f9013ef3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609184"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに場所を取得または設定する

Office JavaScript API には、ユーザーが作成している予定の場所を管理するためのプロパティとメソッドが用意されています。 現時点では、予定の場所を提供するプロパティは2つあります。

- [アイテムの場所](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所の取得と設定を可能にする基本 API。
- [enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所を取得および設定できる拡張 API。また、[場所の種類](/javascript/api/outlook/office.mailboxenums.locationtype)を指定することもできます。 この型は `LocationType.Custom` 、を使用して場所を設定する場合に使用し `item.location` ます。

次の表に、使用可能な場所の Api とモード (つまり、作成または読み取り) を示します。

| API | 適用可能な予定モード |
|---|---|
| [アイテムの場所](/javascript/api/outlook/office.appointmentread#location) | 出席者/閲覧 |
| [getAsync](/javascript/api/outlook/office.location#getasync-options--callback-) | 開催者/新規作成 |
| [item.location.setAsync](/javascript/api/outlook/office.location#setasync-location--options--callback-) | 開催者/新規作成 |
| [enhancedLocation。 getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | 開催者/新規作成<br>出席者/閲覧 |
| [enhancedLocation。 addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | 開催者/新規作成 |
| [enhancedLocation。 removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | 開催者/新規作成 |

アドインの作成にのみ使用できるメソッドを使用するには、アドインマニフェストを構成して、オーガナイザー/新規作成モードでアドインをアクティブにします。 詳細については、「[新規フォーム用の Outlook アドインを作成](compose-scenario.md)する」を参照してください。

## <a name="use-the-enhancedlocation-api"></a>API を使用する `enhancedLocation`

API を使用し `enhancedLocation` て、予定の場所を取得および設定できます。 Location フィールドには複数の場所がサポートされており、それぞれの場所について、表示名、種類、および会議室の電子メールアドレスを設定できます (該当する場合)。 サポートされる場所の種類については、 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)を参照してください。

### <a name="add-location"></a>場所の追加

次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)で[addasync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-)を呼び出すことによって場所を追加する方法を示しています。

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>場所を取得する

次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation)で[getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-)を呼び出すことによって場所を取得する方法を示しています。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)で[removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-)を呼び出すことによって場所を削除する方法を示しています。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>API を使用する `location`

API を使用し `location` て、予定の場所を取得および設定できます。

### <a name="get-the-location"></a>場所を取得する

ここでは、ユーザーが新規作成している予定の配置場所を取得し、それを表示するコード サンプルを示します。

`item.location.getAsync` を使用するためには、非同期呼び出しの状態と結果を確認するコールバック メソッドを提供します。 オプション パラメーターである `asyncContext` を通して、コールバック メソッドに必要な引数を提供できます。 コールバックの出力パラメーターを使用して、状態、結果、およびエラーを取得でき `asyncResult` ます。 非同期コールが成功した場合、[AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティを使用して、配置場所を文字列として取得することができます。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

`item.location.setAsync` を使用するには、data パラメーターに最大 255 文字までの文字列を指定します。 オプションとして、`asyncContext` パラメーターで、コールバック メソッドとそれに必要な引数を提供することができます。 コールバックの出力パラメーターで、状態、結果、およびエラーメッセージを確認する必要があり `asyncResult` ます。 非同期呼び出しが成功した場合、`setAsync` はそのアイテムの既存の配置場所を上書きし、指定した配置場所をプレーンテキストとして挿入します。

> [!NOTE]
> 区切り文字としてセミコロンを使用して、複数の場所を設定できます (たとえば、「会議室 A;」など)。会議室 B ')。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
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
