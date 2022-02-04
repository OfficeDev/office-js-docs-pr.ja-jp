---
title: アドインで予定の場所を取得または設定する
description: Outlook アドインで予定の場所を取得または設定する方法について説明します。
ms.date: 10/31/2019
ms.localizationpriority: medium
---

# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに場所を取得または設定する

JavaScript API Officeには、ユーザーが作成している予定の場所を管理するためのプロパティとメソッドが提供されています。 現在、予定の場所を提供する 2 つのプロパティがあります。

- [item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所を取得および設定できる基本 API。
- [item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所を取得および設定できる拡張 API で、場所の種類の指定が [含まれます](/javascript/api/outlook/office.mailboxenums.locationtype)。 この型は、 `LocationType.Custom` を使用して場所を設定した場合です `item.location`。

次の表に、場所 API と、使用可能なモード (新規作成または読み取り) の一覧を示します。

| API | 適用可能な予定モード |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | 出席者/読み取り |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | オーガナイザー/作成 |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | オーガナイザー/作成 |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | オーガナイザー/作成、<br>出席者/読み取り |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | オーガナイザー/作成 |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | オーガナイザー/作成 |

アドインの作成にのみ使用できるメソッドを使用するには、オーガナイザー/作成モードでアドインをアクティブ化するアドイン マニフェストを構成します。 詳細[についてはOutlook作成](compose-scenario.md)フォームのアドインを作成するを参照してください。

## <a name="use-the-enhancedlocation-api"></a>API を使用 `enhancedLocation` する

API を使用して、 `enhancedLocation` 予定の場所を取得および設定できます。 場所フィールドは複数の場所をサポートし、場所ごとに表示名、種類、および会議室の電子メール アドレスを設定できます (該当する場合)。 サポートされている場所の種類については、「 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) 」を参照してください。

### <a name="add-location"></a>場所の追加

次の例は、[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) で [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) を呼び出して場所を追加する方法を示しています。

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

次の例は、mailbox.item.enhancedLocation で [getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) を呼び出して場所 [を取得する方法を示しています](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member)。

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

次の例は、 [mailbox.item.enhancedLocation で removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) を呼び出して場所 [を削除する方法を示しています](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member)。

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

## <a name="use-the-location-api"></a>API を使用 `location` する

API を使用して、 `location` 予定の場所を取得および設定できます。

### <a name="get-the-location"></a>場所を取得する

ここでは、ユーザーが新規作成している予定の配置場所を取得し、それを表示するコード サンプルを示します。

`item.location.getAsync` を使用するためには、非同期呼び出しの状態と結果を確認するコールバック メソッドを提供します。 オプション パラメーターである `asyncContext` を通して、コールバック メソッドに必要な引数を提供できます。 コールバックの出力パラメーターを使用して、状態、結果、およびエラー `asyncResult` を取得できます。 非同期コールが成功した場合、[AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) プロパティを使用して、配置場所を文字列として取得することができます。

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

`item.location.setAsync` を使用するには、data パラメーターに最大 255 文字までの文字列を指定します。 オプションとして、`asyncContext` パラメーターで、コールバック メソッドとそれに必要な引数を提供することができます。 コールバックの出力パラメーターで、状態、結果 `asyncResult` 、およびエラー メッセージを確認する必要があります。 非同期呼び出しが成功した場合、`setAsync` はそのアイテムの既存の配置場所を上書きし、指定した配置場所をプレーンテキストとして挿入します。

> [!NOTE]
> 区切り記号としてセミコロンを使用して複数の場所を設定できます (例: '会議室 A;会議室 B')。

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

- [最初のアドインOutlook作成する](../quickstarts/outlook-quickstart.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
