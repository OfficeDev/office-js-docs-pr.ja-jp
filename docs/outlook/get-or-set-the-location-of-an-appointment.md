---
title: アドインで予定の場所を取得または設定する
description: Outlook アドインで予定の場所を取得または設定する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: cc412da5dd64d8e908b86a81b847f6479dbd4a34
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324969"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="7c3cc-103">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-103">Get or set the location when composing an appointment in Outlook</span></span>

<span data-ttu-id="7c3cc-104">Office JavaScript API には、ユーザーが作成している予定の場所を管理するためのプロパティとメソッドが用意されています。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-104">The Office JavaScript API provides properties and methods to manage the location of an appointment that the user is composing.</span></span> <span data-ttu-id="7c3cc-105">現時点では、予定の場所を提供するプロパティは2つあります。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-105">Currently, there are two properties that provide an appointment's location:</span></span>

- <span data-ttu-id="7c3cc-106">[アイテムの場所](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所の取得と設定を可能にする基本 API。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-106">[item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.</span></span>
- <span data-ttu-id="7c3cc-107">[enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): 場所を取得および設定できる拡張 API。また、[場所の種類](/javascript/api/outlook/office.mailboxenums.locationtype)を指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-107">[item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype).</span></span> <span data-ttu-id="7c3cc-108">この型は`LocationType.Custom` 、を使用して場所`item.location`を設定する場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-108">The type is `LocationType.Custom` if you set the location using `item.location`.</span></span>

<span data-ttu-id="7c3cc-109">次の表に、使用可能な場所の Api とモード (つまり、作成または読み取り) を示します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-109">The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.</span></span>

| <span data-ttu-id="7c3cc-110">API</span><span class="sxs-lookup"><span data-stu-id="7c3cc-110">API</span></span> | <span data-ttu-id="7c3cc-111">適用可能な予定モード</span><span class="sxs-lookup"><span data-stu-id="7c3cc-111">Applicable appointment modes</span></span> |
|---|---|
| [<span data-ttu-id="7c3cc-112">アイテムの場所</span><span class="sxs-lookup"><span data-stu-id="7c3cc-112">item.location</span></span>](/javascript/api/outlook/office.appointmentread#location) | <span data-ttu-id="7c3cc-113">出席者/閲覧</span><span class="sxs-lookup"><span data-stu-id="7c3cc-113">Attendee/Read</span></span> |
| [<span data-ttu-id="7c3cc-114">getAsync</span><span class="sxs-lookup"><span data-stu-id="7c3cc-114">item.location.getAsync</span></span>](/javascript/api/outlook/office.location#getasync-options--callback-) | <span data-ttu-id="7c3cc-115">開催者/新規作成</span><span class="sxs-lookup"><span data-stu-id="7c3cc-115">Organizer/Compose</span></span> |
| [<span data-ttu-id="7c3cc-116">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="7c3cc-116">item.location.setAsync</span></span>](/javascript/api/outlook/office.location#setasync-location--options--callback-) | <span data-ttu-id="7c3cc-117">開催者/新規作成</span><span class="sxs-lookup"><span data-stu-id="7c3cc-117">Organizer/Compose</span></span> |
| [<span data-ttu-id="7c3cc-118">enhancedLocation。 getAsync</span><span class="sxs-lookup"><span data-stu-id="7c3cc-118">item.enhancedLocation.getAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | <span data-ttu-id="7c3cc-119">開催者/新規作成</span><span class="sxs-lookup"><span data-stu-id="7c3cc-119">Organizer/Compose,</span></span><br><span data-ttu-id="7c3cc-120">出席者/閲覧</span><span class="sxs-lookup"><span data-stu-id="7c3cc-120">Attendee/Read</span></span> |
| [<span data-ttu-id="7c3cc-121">enhancedLocation。 addAsync</span><span class="sxs-lookup"><span data-stu-id="7c3cc-121">item.enhancedLocation.addAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | <span data-ttu-id="7c3cc-122">開催者/新規作成</span><span class="sxs-lookup"><span data-stu-id="7c3cc-122">Organizer/Compose</span></span> |
| [<span data-ttu-id="7c3cc-123">enhancedLocation。 removeAsync</span><span class="sxs-lookup"><span data-stu-id="7c3cc-123">item.enhancedLocation.removeAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | <span data-ttu-id="7c3cc-124">開催者/新規作成</span><span class="sxs-lookup"><span data-stu-id="7c3cc-124">Organizer/Compose</span></span> |

<span data-ttu-id="7c3cc-125">アドインの作成にのみ使用できるメソッドを使用するには、アドインマニフェストを構成して、オーガナイザー/新規作成モードでアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-125">To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode.</span></span> <span data-ttu-id="7c3cc-126">詳細については、「[新規フォーム用の Outlook アドインを作成](compose-scenario.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-126">See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.</span></span>

## <a name="use-the-enhancedlocation-api"></a><span data-ttu-id="7c3cc-127">`enhancedLocation` API を使用する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-127">Use the `enhancedLocation` API</span></span>

<span data-ttu-id="7c3cc-128">`enhancedLocation` API を使用して、予定の場所を取得および設定できます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-128">You can use the `enhancedLocation` API to get and set an appointment's location.</span></span> <span data-ttu-id="7c3cc-129">Location フィールドには複数の場所がサポートされており、それぞれの場所について、表示名、種類、および会議室の電子メールアドレスを設定できます (該当する場合)。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-129">The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable).</span></span> <span data-ttu-id="7c3cc-130">サポートされる場所の種類については、 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-130">See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.</span></span>

### <a name="add-location"></a><span data-ttu-id="7c3cc-131">場所の追加</span><span class="sxs-lookup"><span data-stu-id="7c3cc-131">Add location</span></span>

<span data-ttu-id="7c3cc-132">次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)で[addasync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-)を呼び出すことによって場所を追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-132">The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

### <a name="get-location"></a><span data-ttu-id="7c3cc-133">場所を取得する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-133">Get location</span></span>

<span data-ttu-id="7c3cc-134">次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation)で[getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-)を呼び出すことによって場所を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-134">The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span></span>

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

### <a name="remove-location"></a><span data-ttu-id="7c3cc-135">場所を削除する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-135">Remove location</span></span>

<span data-ttu-id="7c3cc-136">次の例は、 [enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)で[removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-)を呼び出すことによって場所を削除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-136">The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

## <a name="use-the-location-api"></a><span data-ttu-id="7c3cc-137">`location` API を使用する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-137">Use the `location` API</span></span>

<span data-ttu-id="7c3cc-138">`location` API を使用して、予定の場所を取得および設定できます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-138">You can use the `location` API to get and set an appointment's location.</span></span>

### <a name="get-the-location"></a><span data-ttu-id="7c3cc-139">場所を取得する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-139">Get the location</span></span>

<span data-ttu-id="7c3cc-140">ここでは、ユーザーが新規作成している予定の配置場所を取得し、それを表示するコード サンプルを示します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-140">This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.</span></span>

<span data-ttu-id="7c3cc-141">`item.location.getAsync` を使用するためには、非同期呼び出しの状態と結果を確認するコールバック メソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-141">To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call.</span></span> <span data-ttu-id="7c3cc-142">オプション パラメーターである `asyncContext` を通して、コールバック メソッドに必要な引数を提供できます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-142">You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter.</span></span> <span data-ttu-id="7c3cc-143">コールバックの出力パラメーター `asyncResult`を使用して、状態、結果、およびエラーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-143">You can obtain status, results, and any error using the output parameter `asyncResult` of the callback.</span></span> <span data-ttu-id="7c3cc-144">非同期コールが成功した場合、[AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティを使用して、配置場所を文字列として取得することができます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-144">If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>

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

### <a name="set-the-location"></a><span data-ttu-id="7c3cc-145">場所を設定する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-145">Set the location</span></span>

<span data-ttu-id="7c3cc-146">ここでは、ユーザーが新規作成している予定の配置場所を設定するコード サンプルを示します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-146">This section shows a code sample that sets the location of the appointment that the user is composing.</span></span>

<span data-ttu-id="7c3cc-147">`item.location.setAsync` を使用するには、data パラメーターに最大 255 文字までの文字列を指定します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-147">To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter.</span></span> <span data-ttu-id="7c3cc-148">オプションとして、`asyncContext` パラメーターで、コールバック メソッドとそれに必要な引数を提供することができます。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-148">Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter.</span></span> <span data-ttu-id="7c3cc-149">コールバックの`asyncResult`出力パラメーターで、状態、結果、およびエラーメッセージを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-149">You should check the status, result, and any error message in the `asyncResult` output parameter of the callback.</span></span> <span data-ttu-id="7c3cc-150">非同期呼び出しが成功した場合、`setAsync` はそのアイテムの既存の配置場所を上書きし、指定した配置場所をプレーンテキストとして挿入します。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-150">If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.</span></span>

> [!NOTE]
> <span data-ttu-id="7c3cc-151">区切り文字としてセミコロンを使用して、複数の場所を設定できます (たとえば、「会議室 A;」など)。会議室 B ')。</span><span class="sxs-lookup"><span data-stu-id="7c3cc-151">You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7c3cc-152">関連項目</span><span class="sxs-lookup"><span data-stu-id="7c3cc-152">See also</span></span>

- [<span data-ttu-id="7c3cc-153">最初の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="7c3cc-153">Create your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="7c3cc-154">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="7c3cc-154">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
